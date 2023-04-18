using AssetsTools.NET;
using AssetsTools.NET.Extra;
using Mono.Cecil;
using Mono.Cecil.Cil;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace EngageXml
{
    class Program
    {
        static AssetsManager AM;
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                if (File.Exists("Config.xml"))
                {
                    EConfig conf = new EConfig("Config.xml");
                    foreach ( var file in conf.FilePatches)
                    {
                        string source = file.Attribute("source").Value;
                        string target = file.Attribute("target").Value;
                        if (source == "" || target == "") continue;

                        Console.Write($"Processing ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write(source);
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.Write("->");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write(target);
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine();

                        switch (Path.GetExtension(source)) {
                            case ".xml":
                                BundleIO.InsertAsset(File.ReadAllBytes(source), target);
                                break;
                            case ".xlsx":
                                BundleIO.InsertAsset(EXML.FromXlsx(source).ToBinary(), target);
                                break;
                            case ".csv":
                                byte[] data = BundleIO.ReadBundleAssetData(target);
                                MSBT msbt = new MSBT(data);
                                msbt.UpdateWithCSV(source);
                                BundleIO.InsertAsset(msbt.Binarize(), target);
                                break;
                            default:
                                Console.WriteLine($"Skipped file {source}, whose format is not supported, ");
                                break;
                        }
                    }

                    foreach ( var patch in conf.ParamPatches)
                    {
                        string target_file = patch.Attribute("target_file").Value;
                        string target_path = patch.Attribute("target_path").Value;
                        string[] ids = patch.Attribute("IDAttributes").Value.Split(';');
                        if (string.IsNullOrEmpty(target_file) || string.IsNullOrEmpty(target_path)) continue;

                        Console.Write($"Processing ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write(target_file);
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine();

                        EXML exml = new EXML(BundleIO.ReadBundleAssetData(target_file));
                        XElement container = exml.GetElementByPath(target_path);
                        foreach( var param in patch.Elements())
                        {
                            Console.Write($"\tWriting ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write(param);
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine();
                            if (ids.Length > 0)
                            {
                                var ps = from p in container.Elements()
                                where ids.Select(id => p.Attribute(id).Value).SequenceEqual(ids.Select(id => param.Attribute(id).Value))
                                select p;
                                if (ps.Count() > 0)
                                {
                                    foreach (var attr in ps.First().Attributes())
                                    {
                                        attr.Value = param.Attribute(attr.Name).Value;
                                    }
                                }
                                else { container.Add(param); }
                            } else{ container.Add(param);}
                        }
                        BundleIO.InsertAsset(exml.ToBinary(), target_file);
                    }
                    Console.WriteLine("Patch Complete! Press any key to exit");
                    Console.ReadKey();
                    return;
                } else
                {
                    throw new Exception("Config file not found");
                }
            }
            string arg1 = args[0];
            AM = new AssetsManager();
            if (arg1.StartsWith("-"))
            {
                switch(arg1) {
                    case "-in":
                        string dataPath = args[1];
                        byte[] data;
                        if (dataPath.EndsWith(".csv"))
                        {
                            data = UpdateMSBTWithCSV(dataPath, args[2]);
                        } else if (dataPath.EndsWith(".xlsx"))
                        {
                            data = Xlsx2Xml(dataPath);
                        } else
                        {
                            data = File.ReadAllBytes(dataPath);
                        }
                        InsertAsset(data, args[2]);
                        break;
                    case "-out":
                        if (args[1] == "-xlsx")
                        {
                            ExtractAsset(args[2], true);
                        } else
                        {
                            ExtractAsset(args[1], false);
                        }
                        break;
                    case "-update":
                        if(args[1] == "-o")
                        {
                            InsertAsset(ModUpdate(args[2], args[3], true, args.Skip(4).ToArray()), args[3]);
                        } else
                        {
                            InsertAsset(ModUpdate(args[1], args[2], false, args.Skip(3).ToArray()), args[2]);
                        }
                        break;
                    default:
                        throw new Exception($"Invalid option: {arg1}");
                }

            } else
            {
                // conversion between xml and xlsx
                string path = arg1;
                if (!File.Exists(path)) throw new Exception("Error: File not found!");
                if (path.EndsWith(".xml"))
                {
                    File.WriteAllBytes(path + ".xlsx", Xml2Xlsx(path));
                    
                }
                else if (path.EndsWith(".xlsx"))
                {
                    File.WriteAllBytes(Path.ChangeExtension(path, null), Xlsx2Xml(path));
                }
                else if (path.EndsWith(".bundle"))
                {
                    ExtractAsset(path, false);
                } else
                {
                    new Exception("Error: File format not supported!");
                }
            }

        }

        static byte[] Xml2Xlsx(byte[] data)
        {
            XmlDocument xml = new XmlDocument();
            string xmlText = Encoding.UTF8.GetString(data);
            xmlText = xmlText.Substring(1, xmlText.Length - 1);
            xml.LoadXml(xmlText);
            return Xml2XlsxReal(xml);
        }

        static byte[] Xml2Xlsx(string path)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(path);
            return Xml2XlsxReal(xml);
        }

        static byte[] Xml2XlsxReal(XmlDocument xml)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pkg = new ExcelPackage();
            XmlNode bookNode = xml.DocumentElement.SelectSingleNode("/Book");
            
            foreach (XmlNode sheetNode in bookNode.ChildNodes)
            {
                List<string> attrNames = new List<string>();
                int row;
                int col;
                var headerSheet = pkg.Workbook.Worksheets.Add(sheetNode.Attributes["Name"].Value + "Header");
                XmlNode headerNode = sheetNode.SelectSingleNode("Header");
                col = 1;
                foreach ( XmlAttribute attr in headerNode.FirstChild.Attributes)
                {
                    attrNames.Add(attr.Name);
                    headerSheet.Cells[1, col].Value = attr.Name;
                    col++;
                }
                row = 2;
                foreach (XmlNode paramNode in headerNode.ChildNodes)
                {
                    col = 1;
                    foreach(string name in attrNames)
                    {
                        headerSheet.Cells[row, col].Value = paramNode.Attributes[name].Value;
                        col++;
                    }
                    row++;
                }

                attrNames.Clear();
                var dataSheet = pkg.Workbook.Worksheets.Add(sheetNode.Attributes["Name"].Value);
                XmlNode dataNode = sheetNode.SelectSingleNode("Data");
                col = 1;
                foreach (XmlAttribute attr in dataNode.FirstChild.Attributes)
                {
                    attrNames.Add(attr.Name);
                    dataSheet.Cells[1, col].Value = attr.Name;
                    col++;
                }
                row = 2;
                foreach (XmlNode paramNode in dataNode.ChildNodes)
                {
                    col = 1;
                    foreach (string name in attrNames)
                    {
                        dataSheet.Cells[row, col].Value = paramNode.Attributes[name].Value;
                        col++;
                    }
                    row++;
                }
            }


            byte[] xlsxBytes = pkg.GetAsByteArray();
            pkg.Dispose();
            return xlsxBytes;
        }

        static byte[] Xlsx2Xml(string path)
        {
            FileInfo file = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] xmlBytes;

            using ( ExcelPackage pkg = new ExcelPackage(file))
            {
                var sheets = pkg.Workbook.Worksheets
                    .Where(sht => !sht.Name.EndsWith("Header"))
                    .Where(sht => !sht.Name.StartsWith("#"));
                var i = pkg.Workbook.Worksheets.Count();
                // XML 文件生成
                XmlDocument xml = new XmlDocument();
                xml.AppendChild(xml.CreateXmlDeclaration("1.0", "utf-8", null));
                
                // 创建 Book 根节点，指定 Count 属性为 sheet 数量 
                XmlElement nodeBook = xml.CreateElement("Book");
                nodeBook.SetAttribute("Count", sheets.Count().ToString());
                xml.AppendChild(nodeBook);

                foreach (var sheet in sheets)
                {
                    // Book>Sheet 节点
                    XmlElement nodeSheet = xml.CreateElement("Sheet");
                    nodeSheet.SetAttribute("Name", sheet.Name);
                    nodeBook.AppendChild(nodeSheet);
                    List<string> paramAttrs = new List<string>();

                    // 写入 Book>Sheet>Header
                    var hsht = pkg.Workbook.Worksheets[sheet.Name + "Header"];
                    ExcelCellAddress start, end;
                    if (hsht != null)
                    {
                        start = hsht.Dimension.Start;
                        end = hsht.Dimension.End;
                        XmlElement nodeHeader = xml.CreateElement("Header");
                        nodeSheet.AppendChild(nodeHeader);

                        for (int col = start.Column; col <= end.Column; col++)
                        {
                            paramAttrs.Add(hsht.Cells[1, col].Text);
                        }
                        for (int row = start.Row + 1; row <= end.Row; row++)
                        {
                            XmlElement nodeParam = xml.CreateElement("Param");
                            nodeHeader.AppendChild(nodeParam);
                            for (int col = start.Column; col <= end.Column; col++)
                            {
                                nodeParam.SetAttribute(paramAttrs[col - 1], hsht.Cells[row, col].Text);
                            }
                        }
                    }

                    // 写入 Book>Sheet>Data
                    start = sheet.Dimension.Start;
                    end = sheet.Dimension.End;
                    nodeSheet.SetAttribute("Count", (end.Row - 1).ToString());
                    paramAttrs.Clear();
                    XmlElement nodeData = xml.CreateElement("Data");
                    nodeSheet.AppendChild(nodeData);

                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        paramAttrs.Add(sheet.Cells[1, col].Text);
                    }
                    for (int row = start.Row + 1; row <= end.Row; row++)
                    {
                        XmlElement nodeParam = xml.CreateElement("Param");
                        nodeData.AppendChild(nodeParam);
                        for (int col = start.Column; col <= end.Column; col++)
                        {
                            nodeParam.SetAttribute(paramAttrs[col - 1], sheet.Cells[row, col].Text);
                        }
                    }
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    xml.Save(ms);
                    xmlBytes = ms.ToArray();
                }
            }

            return xmlBytes;
        }

        static byte[] ModUpdate(string source_path, string target_path, bool overwrite = true, params string[] ids)
        {
            byte[] updatedBytes;
            XDocument src;
            XDocument tgt;

            string tgtText = Encoding.UTF8.GetString(ReadBundleAssetData(target_path));
            tgtText = tgtText.Substring(1, tgtText.Length - 1);
            tgt = XDocument.Parse(tgtText);

            if (source_path.EndsWith(".xml.bundle"))
            {
                string srcText = Encoding.UTF8.GetString(ReadBundleAssetData(source_path));
                srcText = srcText.Substring(1, srcText.Length - 1);
                src = XDocument.Parse(srcText);
            } else if (source_path.EndsWith(".xml"))
            {
                src = XDocument.Load(source_path);
            } else if (source_path.EndsWith(".xlsx"))
            {
                string srcText = Encoding.UTF8.GetString(Xlsx2Xml(source_path));
                srcText = srcText.Substring(1, srcText.Length - 1);
                src = XDocument.Parse(srcText);
            } else
            {
                throw new Exception("source file not supported");
            }
            Dictionary<string, string> idCache = new Dictionary<string, string>();
            foreach (string id in ids) { idCache.Add(id, ""); }

            foreach(XElement srcSheet in src.Root.Elements("Sheet"))
            {
                var tgtSheet = tgt.Root.Elements("Sheet").Where(sheet=>sheet.Attribute("Name").Value==srcSheet.Attribute("Name").Value).FirstOrDefault();
                if (tgtSheet == null) continue;

                Dictionary<string, XElement> dict = new Dictionary<string, XElement>();
                foreach (XElement param in tgtSheet.Element("Data").Elements())
                {

                    dict.Add(string.Join(",", ids.Select(id => {
                        if (param.Attribute(id).Value == "")
                        {
                            return idCache[id];
                        } else
                        {
                            idCache[id] = param.Attribute(id).Value;
                            return param.Attribute(id).Value;
                        }
                    })), param);
                }

                ids.Select(id => idCache[id] = "");
                foreach (XElement param in srcSheet.Element("Data").Elements())
                {
                    string key = string.Join(",", ids.Select(id =>
                    {
                        if (param.Attribute(id).Value == "")
                        {
                            return idCache[id];
                        }
                        else
                        {
                            idCache[id] = param.Attribute(id).Value;
                            return param.Attribute(id).Value;
                        }
                    }));
                    if (dict.ContainsKey(key))
                    {
                        if (overwrite)
                        {
                            dict[key].ReplaceAttributes(param.Attributes());
                        }
                    }
                    else
                    {
                        tgtSheet.Element("Data").Add(param);
                    }
                }

            }

            using (MemoryStream ms = new MemoryStream())
            {
                tgt.Save(ms);
                updatedBytes = ms.ToArray();
            }

            AM.UnloadAllBundleFiles();
            return updatedBytes;
        }

        static byte[]  ReadBundleAssetData(string bundlePath)
        {
            var bun = AM.LoadBundleFile(bundlePath);

            //load first asset from bundle
            var inst = AM.LoadAssetsFileFromBundle(bun, 0);
            if (!inst.file.typeTree.hasTypeTree)
                AM.LoadClassDatabaseFromPackage(inst.file.typeTree.unityVersion);
            var inf = inst.table.assetFileInfo[0].index == 1 ? inst.table.assetFileInfo[1] : inst.table.assetFileInfo[0];
            var baseField = AM.GetTypeInstance(inst, inf).GetBaseField();
            byte[] data = baseField.Get("m_Script").GetValue().AsStringBytes();

            AM.UnloadAll();
            return data;
        }

        static void InsertAsset(byte[] data, string bundlePath)
        {
            var bun = AM.LoadBundleFile(bundlePath);

            //load first asset from bundle
            var inst = AM.LoadAssetsFileFromBundle(bun, 0);
            if (!inst.file.typeTree.hasTypeTree)
                AM.LoadClassDatabaseFromPackage(inst.file.typeTree.unityVersion);
            var inf = inst.table.assetFileInfo[0].index == 1 ? inst.table.assetFileInfo[1] : inst.table.assetFileInfo[0];
            var baseField = AM.GetTypeInstance(inst, inf).GetBaseField();
            baseField.Get("m_Script").GetValue().Set(data);

            var newGoBytes = baseField.WriteToByteArray();
            var repl = new AssetsReplacerFromMemory(0, inf.index, (int)inf.curFileType, 0xffff, newGoBytes);

            //write changes to memory
            byte[] newAssetData;
            using (var stream = new MemoryStream())
            using (var writer = new AssetsFileWriter(stream))
            {
                inst.file.Write(writer, 0, new List<AssetsReplacer>() { repl }, 0);
                newAssetData = stream.ToArray();
            }

            //rename this asset name from boring to cool when saving
            var bunRepl = new BundleReplacerFromMemory(inst.name, null, true, newAssetData, -1);
            byte[] newBundleData;
            using (var stream = new MemoryStream())
            using (var bunWriter = new AssetsFileWriter(stream))
            {
                bun.file.Write(bunWriter, new List<BundleReplacer>() { bunRepl });
                newBundleData = stream.ToArray();
            }


            MemoryStream newBundleStream = new MemoryStream(newBundleData);
            bun = AM.LoadBundleFile(newBundleStream, $"{bundlePath}.mod");
            AM.UnloadBundleFile(bundlePath);

            using (var stream = File.OpenWrite(bundlePath))
            using (var writer = new AssetsFileWriter(stream))
            {
                bun.file.Pack(bun.file.reader, writer, AssetBundleCompressionType.LZ4);
            }
        }

        static byte[] UpdateMSBTWithCSV(string csvPath, string bundlePath)
        {
            byte[] data = ReadBundleAssetData(bundlePath);
            MSBT msbt = new MSBT(data);
            using (var fs = new FileStream(csvPath, FileMode.Open))
            using (var sr = new StreamReader(fs))
            {
                string[] kv;
                string line, key, value;
                line = sr.ReadLine();
                while (line != null) 
                {
                    line = line.Replace("\\n", "\n");
                    kv = line.Split(',');
                    key = kv[0].Trim();
                    value = kv[1].Trim();
                    if (key.Length <= MSBT.LabelMaxLength && Regex.IsMatch(key, MSBT.LabelFilter))
                    {
                        Label lbl = msbt.HasLabel(key);
                        if (lbl == null) lbl = msbt.AddLabel(key);
                        if (value == "")
                        {
                            msbt.RemoveLabel(lbl);
                        }
                        else
                        {
                            IEntry ent = msbt.TXT2.Strings[(int)lbl.Index];
                            ent.Value = msbt.FileEncoding.GetBytes(value.Replace("\r\n", "\n") + "\0");
                        }
                    } else
                    {
                        throw new Exception("Invalid Label!");
                    }
                    line = sr.ReadLine();
                }
            }
            AM.UnloadBundleFile(bundlePath);
            return msbt.Binarize();
        }

        static void ExtractAsset(string bundlePath, bool toXlsx) 
        {
            byte[] data = ReadBundleAssetData(bundlePath);

            if (!toXlsx)
            {
                    File.WriteAllBytes(Path.ChangeExtension(bundlePath, null), data);
            } else
            {
                File.WriteAllBytes(Path.ChangeExtension(bundlePath, ".xlsx"), Xml2Xlsx(data));
            }
        }
    }
}
