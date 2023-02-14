using AssetsTools.NET;
using AssetsTools.NET.Extra;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace EngageXml
{
    class Program
    {
        static AssetsManager AM;
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                throw new Exception("Please excute this program with arguments.");
            }
            string arg1 = args[0];
            AM = new AssetsManager();
            if (arg1.StartsWith("-"))
            {
                switch(arg1) {
                    case "-in":
                        string dataPath = args[1];
                        byte[] data = dataPath.EndsWith(".xlsx") ? Xlsx2Xml(dataPath): File.ReadAllBytes(dataPath);
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
                            InsertAsset(ModUpdate(args[2], args[3], true), args[3]);
                        } else
                        {
                            InsertAsset(ModUpdate(args[1], args[2], false), args[2]);
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
                    var start = hsht.Dimension.Start;
                    var end = hsht.Dimension.End;
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

        static byte[] ModUpdate(string source_path, string target_path, bool overwrite = true)
        {
            byte[] updatedBytes;
            XmlDocument source_xml = new XmlDocument();
            XmlDocument target_xml = new XmlDocument();
            string xmlText = Encoding.UTF8.GetString(ReadBundleAssetData(target_path));
            xmlText = xmlText.Substring(1, xmlText.Length - 1);
            target_xml.LoadXml(xmlText);
            if (source_path.EndsWith(".xml.bundle"))
            {
                xmlText = Encoding.UTF8.GetString(ReadBundleAssetData(source_path));
                xmlText = xmlText.Substring(1, xmlText.Length - 1);
                source_xml.LoadXml(xmlText);
            } else if (source_path.EndsWith(".xml"))
            {
                source_xml.Load(source_path);
            } else if (source_path.EndsWith(".xlsx"))
            {
                xmlText = Encoding.UTF8.GetString(Xlsx2Xml(source_path));
                xmlText = xmlText.Substring(1, xmlText.Length - 1);
                source_xml.LoadXml(xmlText);
            }
            XmlNode dataSrc = source_xml.DocumentElement.SelectSingleNode("/Book/Sheet/Data");
            XmlNode dataTgt = target_xml.DocumentElement.SelectSingleNode("/Book/Sheet/Data");
            Dictionary<string, XmlNode> dict = new Dictionary<string, XmlNode>();
            foreach (XmlNode param in dataTgt.ChildNodes)
            {
                dict.Add(param.Attributes[1].Value, param);
            }
            foreach (XmlNode param in dataSrc.ChildNodes)
            {
                if (dict.ContainsKey(param.Attributes[1].Value))
                {
                    if (overwrite)
                    {
                        foreach (XmlAttribute attr in dict[param.Attributes[1].Value].Attributes)
                        {
                            attr.Value = param.Attributes[attr.Name].Value;
                        }
                    }
                } else
                {
                    dataTgt.AppendChild(target_xml.ImportNode(param, false));
                } 
            }

            using (MemoryStream ms = new MemoryStream())
            {
                target_xml.Save(ms);
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
