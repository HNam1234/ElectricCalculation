using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class ExcelExportService
    {
        public static void ExportToFile(string templatePath, string outputPath, IEnumerable<Customer> readings)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
            {
                throw new ArgumentException("Template path is required.", nameof(templatePath));
            }

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("Template Excel file not found.", templatePath);
            }

            var readingList = readings?.ToList() ?? new List<Customer>();
            if (readingList.Count == 0)
            {
                throw new InvalidOperationException("Danh sách dữ liệu trống, không có gì để export.");
            }

            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry == null)
            {
                throw new InvalidOperationException("File Excel template không hợp lệ: thiếu xl/workbook.xml.");
            }

            XDocument workbookDoc;
            using (var workbookStream = workbookEntry.Open())
            {
                workbookDoc = XDocument.Load(workbookStream);
            }
            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets");
            if (sheetsElement == null)
            {
                throw new InvalidOperationException("File Excel template không hợp lệ: không tìm thấy danh sách sheet.");
            }

            var dataSheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), "Data", StringComparison.OrdinalIgnoreCase));

            if (dataSheetElement == null)
            {
                throw new InvalidOperationException("File Excel template không có sheet 'Data'.");
            }

            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);
            var relId = (string?)dataSheetElement.Attribute(relIdAttributeName);
            if (string.IsNullOrWhiteSpace(relId))
            {
                throw new InvalidOperationException("Sheet 'Data' không có liên kết tới nội dung.");
            }

            var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (relEntry == null)
            {
                throw new InvalidOperationException("File Excel template không hợp lệ: thiếu xl/_rels/workbook.xml.rels.");
            }

            XDocument relDoc;
            using (var relStream = relEntry.Open())
            {
                relDoc = XDocument.Load(relStream);
            }
            var relationship = relDoc
                .Root?
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), relId, StringComparison.Ordinal));

            var target = (string?)relationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                throw new InvalidOperationException("Không xác định được đường dẫn nội dung cho sheet 'Data'.");
            }

            var sheetPath = target!.StartsWith("/")
                ? "xl" + target
                : "xl/" + target;

            var sheetEntry = archive.GetEntry(sheetPath);
            if (sheetEntry == null)
            {
                throw new InvalidOperationException($"File Excel template không hợp lệ: thiếu {sheetPath}.");
            }

            XDocument sheetDoc;
            using (var sheetReadStream = sheetEntry.Open())
            {
                sheetDoc = XDocument.Load(sheetReadStream);
            }
            var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData");
            if (sheetDataElement == null)
            {
                throw new InvalidOperationException("Sheet 'Data' không chứa phần tử sheetData.");
            }

            // Remove existing data rows (keep header and title rows, i.e. rows < 5)
            var rowsToRemove = sheetDataElement
                .Elements(mainNs + "row")
                .Where(r =>
                {
                    if (!int.TryParse((string?)r.Attribute("r"), out var rowIndex))
                    {
                        return false;
                    }

                    return rowIndex >= 5;
                })
                .ToList();

            foreach (var row in rowsToRemove)
            {
                row.Remove();
            }

            var rowIndexCounter = 5;
            var stt = 1;

            foreach (var vm in readingList)
            {
                var rowElement = new XElement(mainNs + "row",
                    new XAttribute("r", rowIndexCounter));

                // STT
                rowElement.Add(CreateNumberCell(mainNs, "A", rowIndexCounter, stt));

                // Customer info
                if (!string.IsNullOrWhiteSpace(vm.Name))
                {
                    rowElement.Add(CreateTextCell(mainNs, "B", rowIndexCounter, vm.Name));
                }

                if (!string.IsNullOrWhiteSpace(vm.GroupName))
                {
                    rowElement.Add(CreateTextCell(mainNs, "C", rowIndexCounter, vm.GroupName));
                }

                if (!string.IsNullOrWhiteSpace(vm.Address))
                {
                    rowElement.Add(CreateTextCell(mainNs, "D", rowIndexCounter, vm.Address));
                }

                if (!string.IsNullOrWhiteSpace(vm.Phone))
                {
                    rowElement.Add(CreateTextCell(mainNs, "E", rowIndexCounter, vm.Phone));
                }

                if (!string.IsNullOrWhiteSpace(vm.MeterNumber))
                {
                    rowElement.Add(CreateTextCell(mainNs, "J", rowIndexCounter, vm.MeterNumber));
                }

                // Numeric values
                rowElement.Add(CreateNumberCell(mainNs, "O", rowIndexCounter, vm.CurrentIndex));
                rowElement.Add(CreateNumberCell(mainNs, "P", rowIndexCounter, vm.PreviousIndex));
                rowElement.Add(CreateNumberCell(mainNs, "Q", rowIndexCounter, vm.Multiplier));
                rowElement.Add(CreateNumberCell(mainNs, "S", rowIndexCounter, vm.SubsidizedKwh));
                rowElement.Add(CreateNumberCell(mainNs, "U", rowIndexCounter, vm.UnitPrice));

                // Formulas similar to original file
                var rAddress = $"R{rowIndexCounter}";
                var sAddress = $"S{rowIndexCounter}";
                var tAddress = $"T{rowIndexCounter}";
                var uAddress = $"U{rowIndexCounter}";
                var vAddress = $"V{rowIndexCounter}";
                var oAddress = $"O{rowIndexCounter}";
                var pAddress = $"P{rowIndexCounter}";
                var qAddress = $"Q{rowIndexCounter}";

                // R: Tổng điện năng tiêu thụ = (O - P) * Q
                rowElement.Add(CreateFormulaCell(mainNs, "R", rowIndexCounter, $"({oAddress}-{pAddress})*{qAddress}"));

                // T: Điện năng phải trả = IF((R-S)>0,(R-S),0)
                rowElement.Add(CreateFormulaCell(mainNs, "T", rowIndexCounter, $"IF(({rAddress}-{sAddress})>0,({rAddress}-{sAddress}),0)"));

                // V: Thành tiền = T * U
                rowElement.Add(CreateFormulaCell(mainNs, "V", rowIndexCounter, $"{tAddress}*{uAddress}"));

                sheetDataElement.Add(rowElement);

                rowIndexCounter++;
                stt++;
            }

            using (var sheetStream = sheetEntry.Open())
            {
                sheetStream.SetLength(0);
                sheetDoc.Save(sheetStream);
            }
        }

        private static XElement CreateTextCell(XNamespace ns, string column, int rowIndex, string text)
        {
            var cellReference = column + rowIndex.ToString(CultureInfo.InvariantCulture);

            return new XElement(ns + "c",
                new XAttribute("r", cellReference),
                new XAttribute("t", "inlineStr"),
                new XElement(ns + "is",
                    new XElement(ns + "t", text)));
        }

        private static XElement CreateNumberCell(XNamespace ns, string column, int rowIndex, decimal value)
        {
            var cellReference = column + rowIndex.ToString(CultureInfo.InvariantCulture);

            return new XElement(ns + "c",
                new XAttribute("r", cellReference),
                new XElement(ns + "v", value.ToString(CultureInfo.InvariantCulture)));
        }

        private static XElement CreateNumberCell(XNamespace ns, string column, int rowIndex, int value)
        {
            var cellReference = column + rowIndex.ToString(CultureInfo.InvariantCulture);

            return new XElement(ns + "c",
                new XAttribute("r", cellReference),
                new XElement(ns + "v", value.ToString(CultureInfo.InvariantCulture)));
        }

        private static XElement CreateFormulaCell(XNamespace ns, string column, int rowIndex, string formula)
        {
            var cellReference = column + rowIndex.ToString(CultureInfo.InvariantCulture);

            return new XElement(ns + "c",
                new XAttribute("r", cellReference),
                new XElement(ns + "f", formula));
        }
    }
}
