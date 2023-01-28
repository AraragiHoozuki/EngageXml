using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EngageXml
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = args[0];
            if (!File.Exists(path)) throw new Exception("Error: File not found!"); 
            if (path.EndsWith(".xml"))
            {
                Xml2Xlsx(path);
            } else if (path.EndsWith(".xlsx"))
            {
                Xlsx2Xml(path);
            } else
            {
                new Exception("Error: File format not supported!");
            }
        }


        static void Xml2Xlsx(string path)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(path);
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



            FileStream fs = File.Create($"{path}.xlsx");
            fs.Close();
            File.WriteAllBytes($"{path}.xlsx", pkg.GetAsByteArray());
            pkg.Dispose();
        }

        static void Xlsx2Xml(string path)
        {
            FileInfo file = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

                    xml.Save(Path.ChangeExtension(path, null));
                }
            }
        }
    }
}
