using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace EngageXml
{
    public static class XExtensions
    {
        public static XElement GetHeader(this XElement sheet)
        {
            return sheet.Element("Header");
        }

        public static XElement GetData(this XElement sheet)
        {
            return sheet.Element("Data");
        }
    }
    internal class EXML
    {
        private XDocument xml;
        public EXML(string text) {
            xml = XDocument.Parse(text);
        }

        public EXML(byte[] data)
        {
            string text = Encoding.UTF8.GetString(data);
            text = text.Substring(1, text.Length - 1);
            xml = XDocument.Parse(text);
        }

        public static EXML FromXlsx(string path)
        {
            FileInfo file = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] xmlBytes;

            using (ExcelPackage pkg = new ExcelPackage(file))
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

            return new EXML(xmlBytes);
        }

        public XElement Book { get { return xml.Element("Book"); } }
        public IEnumerable<XElement> Sheets { get { return Book.Elements("Sheet"); }}

        public XElement GetSheet(string name)
        {
            return Sheets.Where(sheet => sheet.Attribute("Name").Value == name).FirstOrDefault();
        }
        public XElement GetSheet(int index)
        {
            return Sheets.ToArray()[index];
        }

        public XElement GetElementByPath(string path)
        {
            return xml.XPathSelectElement(path);
        }

        public byte[] ToXlsx()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pkg = new ExcelPackage();

            foreach(var sheet in Sheets)
            {
                // Sheet/Header
                var header = sheet.GetHeader();
                var header_xlsx = pkg.Workbook.Worksheets.Add(sheet.Attribute("Name").Value + "Header");
                List<string> attrNames = new List<string>();
                int col = 1;
                foreach (var attr in header.Element("Param").Attributes())
                {
                    attrNames.Add(attr.Name.ToString());
                    header_xlsx.Cells[1, col].Value = attr.Name.ToString();
                    col++;
                }
                int row = 2;
                foreach (var param in header.Elements("Param"))
                {
                    col = 1;
                    foreach (string name in attrNames)
                    {
                        header_xlsx.Cells[row, col].Value = param.Attribute(name).Value;
                        col++;
                    }
                    row++;
                }
                attrNames.Clear();
                // Sheet/Data
                var data = sheet.GetData();
                var data_xlsx = pkg.Workbook.Worksheets.Add(sheet.Attribute("Name").Value);
                col = 1;
                foreach (var attr in data.Element("Param").Attributes())
                {
                    attrNames.Add(attr.Name.ToString());
                    data_xlsx.Cells[1, col].Value = attr.Name.ToString();
                    col++;
                }
                row = 2;
                foreach (var param in data.Elements("Param"))
                {
                    col = 1;
                    foreach (string name in attrNames)
                    {
                        data_xlsx.Cells[row, col].Value = param.Attribute(name).Value;
                        col++;
                    }
                    row++;
                }
            }

            byte[] xlsxBytes = pkg.GetAsByteArray();
            pkg.Dispose();
            return xlsxBytes;
        }

        public byte[] ToBinary()
        {
            byte[] xmlBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                xml.Save(ms);
                xmlBytes = ms.ToArray();
            }
            return xmlBytes;
        }
    }
}
