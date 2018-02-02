using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;

namespace MultipartExcel
{
    internal static class Program
    {
        private static string GenerateExcelId() =>
            $"R{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 16)}";

        public static void Main(string[] args)
        {
            const int chunksCount = 5;
            var watch = Stopwatch.StartNew();
            Console.WriteLine($"Export started. Exporting {chunksCount} chunks of entities!");

            var sheetIds = new Dictionary<int, string>();
            for (var i = 1; i <= chunksCount; i++)
            {
                sheetIds[i] = GenerateExcelId();
            }

            const string serviceDirectory = "D:/test/multipart_excel_files";
            var reportName = Guid.NewGuid();
            var filesPath = $"{serviceDirectory}/0_compressed";
            Directory.CreateDirectory(filesPath);
            var filePath = $"{serviceDirectory}/{reportName}";
            Directory.CreateDirectory(filePath);
            using (var file = File.OpenWrite($"{filePath}/[Content_Types].xml"))
            {
                var xmlDocument = new XmlDocument();
                var xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "utf-8", null);
                xmlDocument.AppendChild(xmlDeclaration);

                var typesElement = xmlDocument.CreateElement("Types");
                typesElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
                xmlDocument.AppendChild(typesElement);

                var defaultExtension1 = xmlDocument.CreateElement("Default");
                defaultExtension1.SetAttribute("Extension", "xml");
                defaultExtension1.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                typesElement.AppendChild(defaultExtension1);

                var defaultExtension2 = xmlDocument.CreateElement("Default");
                defaultExtension2.SetAttribute("Extension", "rels");
                defaultExtension2.SetAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                typesElement.AppendChild(defaultExtension2);

                for (var i = 1; i <= chunksCount; i++)
                {
                    var overridePartName = xmlDocument.CreateElement("Override");
                    overridePartName.SetAttribute("PartName", $"/xl/worksheets/sheet{i}.xml");
                    overridePartName.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    typesElement.AppendChild(overridePartName);
                }

                var bytes = Encoding.UTF8.GetBytes(xmlDocument.OuterXml);
                file.Write(bytes, 0, bytes.Length);
            }

            Directory.CreateDirectory($"{filePath}/_rels");
            using (var file = File.OpenWrite($"{filePath}/_rels/.rels"))
            {
                var xmlDocument = new XmlDocument();
                var xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "utf-8", null);
                xmlDocument.AppendChild(xmlDeclaration);

                var relationshipsElement = xmlDocument.CreateElement("Relationships");
                relationshipsElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
                xmlDocument.AppendChild(relationshipsElement);

                var relationshipElement = xmlDocument.CreateElement("Relationship");
                relationshipElement.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                relationshipElement.SetAttribute("Target", "/xl/workbook.xml");
                relationshipElement.SetAttribute("Id", GenerateExcelId());
                relationshipsElement.AppendChild(relationshipElement);

                var bytes = Encoding.UTF8.GetBytes(xmlDocument.OuterXml);
                file.Write(bytes, 0, bytes.Length);
            }

            Directory.CreateDirectory($"{filePath}/xl");
            using (var file = File.OpenWrite($"{filePath}/xl/workbook.xml"))
            {
                var fileContentBuilder = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                fileContentBuilder.Append("<x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheets>");
                for (var i = 1; i <= chunksCount; i++)
                {
                    fileContentBuilder.Append($"<x:sheet name=\"Display Sheet {i}\" sheetId=\"{i}\" r:id=\"{sheetIds[i]}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");
                }
                fileContentBuilder.Append("</x:sheets></x:workbook>");
                var bytes = Encoding.UTF8.GetBytes(fileContentBuilder.ToString());
                file.Write(bytes, 0, bytes.Length);
            }

            Directory.CreateDirectory($"{filePath}/xl/_rels");
            using (var file = File.OpenWrite($"{filePath}/xl/_rels/workbook.xml.rels"))
            {
                var xmlDocument = new XmlDocument();
                var xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "utf-8", null);
                xmlDocument.AppendChild(xmlDeclaration);

                var relationshipsElement = xmlDocument.CreateElement("Relationships");
                relationshipsElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
                xmlDocument.AppendChild(relationshipsElement);

                for (var i = 1; i <= chunksCount; i++)
                {
                    var relationshipElement = xmlDocument.CreateElement("Relationship");
                    relationshipElement.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    relationshipElement.SetAttribute("Target", $"/xl/worksheets/sheet{i}.xml");
                    relationshipElement.SetAttribute("Id", $"{sheetIds[i]}");
                    relationshipsElement.AppendChild(relationshipElement);
                }

                var bytes = Encoding.UTF8.GetBytes(xmlDocument.OuterXml);
                file.Write(bytes, 0, bytes.Length);
            }

            Directory.CreateDirectory($"{filePath}/xl/worksheets");
            for (var i = 1; i <= chunksCount; i++)
            {
                using (var fs = new FileStream($"{filePath}/xl/worksheets/sheet{i}.xml", FileMode.Append,
                    FileAccess.Write, FileShare.None, 1 << 20, true))
                {
                    const string start = "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData>";
                    var startBytes = Encoding.UTF8.GetBytes(start);
                    fs.Write(startBytes, 0, startBytes.Length);

                    var propertyInfos = typeof(TestData).GetTypeInfo().GetProperties();
                    var headerBuilder = new StringBuilder("<x:row>");
                    foreach (var propertyInfo in propertyInfos)
                    {
                        headerBuilder.Append($"<x:c t=\"str\"><x:v>{propertyInfo.Name}</x:v></x:c>");
                    }
                    headerBuilder.Append("</x:row>");
                    var headerBytes = Encoding.UTF8.GetBytes(headerBuilder.ToString());
                    fs.Write(headerBytes, 0, headerBytes.Length);

                    var testDataChunk = DataProvider.CreateTestDataChunk();
                    foreach (var testData in testDataChunk)
                    {
                        var entityBuilder = new StringBuilder("<x:row>");

                        foreach (var propertyInfo in propertyInfos)
                        {
                            entityBuilder.Append($"<x:c t=\"str\"><x:v>{propertyInfo.GetValue(testData)}</x:v></x:c>");
                        }

                        entityBuilder.Append("</x:row>");

                        var bytes = Encoding.UTF8.GetBytes(entityBuilder.ToString());
                        fs.Write(bytes, 0, bytes.Length);
                    }

                    const string finish = "</x:sheetData></x:worksheet>";
                    var finishBytes = Encoding.UTF8.GetBytes(finish);
                    fs.Write(finishBytes, 0, finishBytes.Length);
                }
            }

            watch.Stop();
            Console.WriteLine($"Export done. Exported {chunksCount} with {DataProvider.ChunkSize} entities each.");
            Console.WriteLine($"Elapsed: {watch.ElapsedMilliseconds}");
            Console.WriteLine();
            Console.WriteLine();

            watch.Restart();
            ZipFile.CreateFromDirectory(filePath, $"{filesPath}/{reportName}.xlsx");

            watch.Stop();
            Console.WriteLine($"Compressing done. Elapsed: {watch.ElapsedMilliseconds}");
        }
    }
}