using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class ExcelExportService
    {
        public static void ExportToFile(
            string templatePath,
            string outputPath,
            IEnumerable<Customer> readings,
            string? periodLabel = null,
            string? issuerName = null)
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
            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);

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

            var relationshipsRoot = relDoc.Root
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: thiếu relationships root.");

            // Remove calcChain to avoid Excel "We found a problem with some content..." warning.
            const string calcChainType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

            var calcChainRelationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Type"), calcChainType, StringComparison.Ordinal));

            if (calcChainRelationship != null)
            {
                var calcTarget = (string?)calcChainRelationship.Attribute("Target");
                calcChainRelationship.Remove();

                if (!string.IsNullOrWhiteSpace(calcTarget))
                {
                    var calcPath = calcTarget.StartsWith("/", StringComparison.Ordinal)
                        ? "xl" + calcTarget
                        : "xl/" + calcTarget;

                    archive.GetEntry(calcPath)?.Delete();
                }

                using var relWriteStream = relEntry.Open();
                relWriteStream.SetLength(0);
                relDoc.Save(relWriteStream);
            }

            var title = BuildSummaryTitle(periodLabel);

            if (!TryExportSummaryLikeSheet(
                    archive,
                    sheetsElement,
                    relationshipsRoot,
                    relIdAttributeName,
                    mainNs,
                    relPackageNs,
                    sheetName: "Data",
                    readings: readingList,
                    title: title))
            {
                throw new InvalidOperationException("File Excel template không có sheet 'Data'.");
            }

            TryExportSummaryLikeSheet(
                archive,
                sheetsElement,
                relationshipsRoot,
                relIdAttributeName,
                mainNs,
                relPackageNs,
                sheetName: "Bảng kê",
                readings: readingList,
                title: title);

            TryExportPrintBookSheet(
                archive,
                sheetsElement,
                relationshipsRoot,
                relIdAttributeName,
                mainNs,
                relPackageNs,
                sheetName: "Ban  in so",
                readings: readingList,
                title: title);

            TryExportSchoolSheets(
                archive,
                sheetsElement,
                relationshipsRoot,
                relIdAttributeName,
                mainNs,
                relPackageNs,
                readingList,
                periodLabel,
                issuerName);
        }

        private static bool TryExportSummaryLikeSheet(
            ZipArchive archive,
            XElement sheetsElement,
            XElement relationshipsRoot,
            XName relIdAttributeName,
            XNamespace mainNs,
            XNamespace relPackageNs,
            string sheetName,
            IReadOnlyList<Customer> readings,
            string? title)
        {
            if (!TryLoadWorksheet(
                    archive,
                    sheetsElement,
                    relationshipsRoot,
                    relIdAttributeName,
                    mainNs,
                    relPackageNs,
                    sheetName,
                    out var sheetEntry,
                    out var sheetDoc,
                    out var sheetDataElement))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(title))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A1", title);
            }

            RemoveRowsFrom(sheetDataElement, mainNs, startRowIndex: 5);

            var rowIndexCounter = 5;
            var fallbackStt = 1;

            foreach (var vm in readings)
            {
                var rowElement = new XElement(mainNs + "row",
                    new XAttribute("r", rowIndexCounter));

                var sttValue = vm.SequenceNumber > 0 ? vm.SequenceNumber : fallbackStt;

                rowElement.Add(CreateNumberCell(mainNs, "A", rowIndexCounter, sttValue));

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

                var householdPhone = string.IsNullOrWhiteSpace(vm.HouseholdPhone)
                    ? vm.Phone
                    : vm.HouseholdPhone;

                if (!string.IsNullOrWhiteSpace(householdPhone))
                {
                    rowElement.Add(CreateTextCell(mainNs, "E", rowIndexCounter, householdPhone));
                }

                if (!string.IsNullOrWhiteSpace(vm.RepresentativeName))
                {
                    rowElement.Add(CreateTextCell(mainNs, "F", rowIndexCounter, vm.RepresentativeName));
                }

                if (!string.IsNullOrWhiteSpace(vm.Phone))
                {
                    rowElement.Add(CreateTextCell(mainNs, "G", rowIndexCounter, vm.Phone));
                }

                if (!string.IsNullOrWhiteSpace(vm.BuildingName))
                {
                    rowElement.Add(CreateTextCell(mainNs, "H", rowIndexCounter, vm.BuildingName));
                }

                if (!string.IsNullOrWhiteSpace(vm.MeterNumber))
                {
                    rowElement.Add(CreateTextCell(mainNs, "J", rowIndexCounter, vm.MeterNumber));
                }

                if (!string.IsNullOrWhiteSpace(vm.Category))
                {
                    rowElement.Add(CreateTextCell(mainNs, "K", rowIndexCounter, vm.Category));
                }

                if (!string.IsNullOrWhiteSpace(vm.Location))
                {
                    rowElement.Add(CreateTextCell(mainNs, "L", rowIndexCounter, vm.Location));
                }

                if (!string.IsNullOrWhiteSpace(vm.Substation))
                {
                    rowElement.Add(CreateTextCell(mainNs, "M", rowIndexCounter, vm.Substation));
                }

                if (!string.IsNullOrWhiteSpace(vm.Page))
                {
                    rowElement.Add(CreateTextCell(mainNs, "N", rowIndexCounter, vm.Page));
                }

                rowElement.Add(CreateNumberCell(mainNs, "O", rowIndexCounter, vm.CurrentIndex));
                rowElement.Add(CreateNumberCell(mainNs, "P", rowIndexCounter, vm.PreviousIndex));
                rowElement.Add(CreateNumberCell(mainNs, "Q", rowIndexCounter, vm.Multiplier));
                rowElement.Add(CreateNumberCell(mainNs, "S", rowIndexCounter, vm.SubsidizedKwh));
                rowElement.Add(CreateNumberCell(mainNs, "U", rowIndexCounter, vm.UnitPrice));

                var rAddress = $"R{rowIndexCounter}";
                var sAddress = $"S{rowIndexCounter}";
                var tAddress = $"T{rowIndexCounter}";
                var uAddress = $"U{rowIndexCounter}";
                var oAddress = $"O{rowIndexCounter}";
                var pAddress = $"P{rowIndexCounter}";
                var qAddress = $"Q{rowIndexCounter}";

                rowElement.Add(CreateFormulaCell(mainNs, "R", rowIndexCounter, $"({oAddress}-{pAddress})*{qAddress}"));
                rowElement.Add(CreateFormulaCell(mainNs, "T", rowIndexCounter, $"IF(({rAddress}-{sAddress})>0,({rAddress}-{sAddress}),0)"));
                rowElement.Add(CreateFormulaCell(mainNs, "V", rowIndexCounter, $"{tAddress}*{uAddress}"));

                if (!string.IsNullOrWhiteSpace(vm.PerformedBy))
                {
                    rowElement.Add(CreateTextCell(mainNs, "W", rowIndexCounter, vm.PerformedBy));
                }

                sheetDataElement.Add(rowElement);

                rowIndexCounter++;
                fallbackStt++;
            }

            using (var sheetStream = sheetEntry.Open())
            {
                sheetStream.SetLength(0);
                sheetDoc.Save(sheetStream);
            }

            return true;
        }

        private static bool TryExportPrintBookSheet(
            ZipArchive archive,
            XElement sheetsElement,
            XElement relationshipsRoot,
            XName relIdAttributeName,
            XNamespace mainNs,
            XNamespace relPackageNs,
            string sheetName,
            IReadOnlyList<Customer> readings,
            string? title)
        {
            if (!TryLoadWorksheet(
                    archive,
                    sheetsElement,
                    relationshipsRoot,
                    relIdAttributeName,
                    mainNs,
                    relPackageNs,
                    sheetName,
                    out var sheetEntry,
                    out var sheetDoc,
                    out var sheetDataElement))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(title))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A1", title);
            }

            RemoveRowsFrom(sheetDataElement, mainNs, startRowIndex: 4);

            var rowIndexCounter = 4;
            var fallbackStt = 1;

            foreach (var vm in readings)
            {
                var rowElement = new XElement(mainNs + "row",
                    new XAttribute("r", rowIndexCounter));

                var sttValue = vm.SequenceNumber > 0 ? vm.SequenceNumber : fallbackStt;
                rowElement.Add(CreateNumberCell(mainNs, "A", rowIndexCounter, sttValue));

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

                var householdPhone = string.IsNullOrWhiteSpace(vm.HouseholdPhone)
                    ? vm.Phone
                    : vm.HouseholdPhone;

                if (!string.IsNullOrWhiteSpace(householdPhone))
                {
                    rowElement.Add(CreateTextCell(mainNs, "E", rowIndexCounter, householdPhone));
                }

                if (!string.IsNullOrWhiteSpace(vm.RepresentativeName))
                {
                    rowElement.Add(CreateTextCell(mainNs, "F", rowIndexCounter, vm.RepresentativeName));
                }

                if (!string.IsNullOrWhiteSpace(vm.Phone))
                {
                    rowElement.Add(CreateTextCell(mainNs, "G", rowIndexCounter, vm.Phone));
                }

                if (!string.IsNullOrWhiteSpace(vm.BuildingName))
                {
                    rowElement.Add(CreateTextCell(mainNs, "H", rowIndexCounter, vm.BuildingName));
                }

                if (!string.IsNullOrWhiteSpace(vm.Page))
                {
                    rowElement.Add(CreateTextCell(mainNs, "I", rowIndexCounter, vm.Page));
                }

                if (!string.IsNullOrWhiteSpace(vm.MeterNumber))
                {
                    rowElement.Add(CreateTextCell(mainNs, "J", rowIndexCounter, vm.MeterNumber));
                }

                if (!string.IsNullOrWhiteSpace(vm.Category))
                {
                    rowElement.Add(CreateTextCell(mainNs, "K", rowIndexCounter, vm.Category));
                }

                if (!string.IsNullOrWhiteSpace(vm.Location))
                {
                    rowElement.Add(CreateTextCell(mainNs, "L", rowIndexCounter, vm.Location));
                }

                if (!string.IsNullOrWhiteSpace(vm.Substation))
                {
                    rowElement.Add(CreateTextCell(mainNs, "M", rowIndexCounter, vm.Substation));
                }

                rowElement.Add(CreateNumberCell(mainNs, "N", rowIndexCounter, vm.Multiplier));
                rowElement.Add(CreateNumberCell(mainNs, "O", rowIndexCounter, vm.CurrentIndex));
                rowElement.Add(CreateNumberCell(mainNs, "P", rowIndexCounter, vm.PreviousIndex));

                sheetDataElement.Add(rowElement);

                rowIndexCounter++;
                fallbackStt++;
            }

            using (var sheetStream = sheetEntry.Open())
            {
                sheetStream.SetLength(0);
                sheetDoc.Save(sheetStream);
            }

            return true;
        }

        private static void TryExportSchoolSheets(
            ZipArchive archive,
            XElement sheetsElement,
            XElement relationshipsRoot,
            XName relIdAttributeName,
            XNamespace mainNs,
            XNamespace relPackageNs,
            IReadOnlyList<Customer> readings,
            string? periodLabel,
            string? issuerName)
        {
            var schoolSheetNames = sheetsElement
                .Elements(mainNs + "sheet")
                .Select(s => (string?)s.Attribute("name"))
                .Where(name => !string.IsNullOrWhiteSpace(name) && name.TrimStart().StartsWith("Trường", StringComparison.OrdinalIgnoreCase))
                .Select(name => name!.Trim())
                .ToList();

            if (schoolSheetNames.Count == 0)
            {
                return;
            }

            var periodText = BuildSchoolPeriodText(periodLabel);
            var dateText = BuildHanoiDateLine(DateTime.Today);

            foreach (var sheetName in schoolSheetNames)
            {
                if (!TryLoadWorksheet(
                        archive,
                        sheetsElement,
                        relationshipsRoot,
                        relIdAttributeName,
                        mainNs,
                        relPackageNs,
                        sheetName,
                        out var sheetEntry,
                        out var sheetDoc,
                        out var sheetDataElement))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(periodText))
                {
                    UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
                }

                UpdateTextCell(sheetDataElement, mainNs, "A5", $"Kính gửi: {sheetName}");
                UpdateTextCell(sheetDataElement, mainNs, "A7", $"Địa chỉ hộ tiêu thụ: {sheetName}.");
                UpdateTextCell(sheetDataElement, mainNs, "A8", $"Đại diện: {sheetName}.");

                var groupCustomers = readings
                    .Where(c => string.Equals(NormalizeSheetKey(c.GroupName), NormalizeSheetKey(sheetName), StringComparison.OrdinalIgnoreCase))
                    .OrderBy(c => c.SequenceNumber > 0 ? c.SequenceNumber : int.MaxValue)
                    .ThenBy(c => c.Name)
                    .ToList();

                const int startRowIndex = 13;
                const int maxRowCount = 23;

                var fallbackStt = 1;

                for (var i = 0; i < maxRowCount; i++)
                {
                    var rowIndex = startRowIndex + i;

                    if (i < groupCustomers.Count)
                    {
                        var customer = groupCustomers[i];
                        var sttValue = customer.SequenceNumber > 0 ? customer.SequenceNumber : fallbackStt++;

                        UpdateNumberCell(sheetDataElement, mainNs, $"A{rowIndex}", sttValue);
                        UpdateNumberCell(sheetDataElement, mainNs, $"B{rowIndex}", customer.CurrentIndex);
                        UpdateNumberCell(sheetDataElement, mainNs, $"C{rowIndex}", customer.PreviousIndex);
                        UpdateNumberCell(sheetDataElement, mainNs, $"D{rowIndex}", customer.Multiplier);
                        UpdateNumberCell(sheetDataElement, mainNs, $"F{rowIndex}", customer.SubsidizedKwh);
                        UpdateNumberCell(sheetDataElement, mainNs, $"G{rowIndex}", customer.UnitPrice);
                        UpdateTextCell(sheetDataElement, mainNs, $"I{rowIndex}", customer.Name ?? string.Empty);
                        UpdateTextCell(sheetDataElement, mainNs, $"J{rowIndex}", customer.MeterNumber ?? string.Empty);

                        continue;
                    }

                    ClearCellValue(sheetDataElement, mainNs, $"A{rowIndex}");
                    ClearCellValue(sheetDataElement, mainNs, $"B{rowIndex}");
                    ClearCellValue(sheetDataElement, mainNs, $"C{rowIndex}");
                    ClearCellValue(sheetDataElement, mainNs, $"D{rowIndex}");
                    ClearCellValue(sheetDataElement, mainNs, $"F{rowIndex}");
                    ClearCellValue(sheetDataElement, mainNs, $"G{rowIndex}");
                    UpdateTextCell(sheetDataElement, mainNs, $"I{rowIndex}", string.Empty);
                    UpdateTextCell(sheetDataElement, mainNs, $"J{rowIndex}", string.Empty);
                }

                var totalAmount = groupCustomers.Take(maxRowCount).Sum(c => c.Amount);
                var amountText = VietnameseNumberTextService.ConvertAmountToText(totalAmount);
                UpdateTextCell(sheetDataElement, mainNs, "A37", $"Bằng chữ: {amountText} ./.");

                UpdateTextCell(sheetDataElement, mainNs, "H39", dateText);

                if (!string.IsNullOrWhiteSpace(issuerName))
                {
                    UpdateTextCell(sheetDataElement, mainNs, "H44", issuerName.Trim());
                }

                using (var sheetStream = sheetEntry.Open())
                {
                    sheetStream.SetLength(0);
                    sheetDoc.Save(sheetStream);
                }
            }
        }

        private static bool TryLoadWorksheet(
            ZipArchive archive,
            XElement sheetsElement,
            XElement relationshipsRoot,
            XName relIdAttributeName,
            XNamespace mainNs,
            XNamespace relPackageNs,
            string sheetName,
            out ZipArchiveEntry sheetEntry,
            out XDocument sheetDoc,
            out XElement sheetDataElement)
        {
            sheetEntry = null!;
            sheetDoc = null!;
            sheetDataElement = null!;

            var sheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheetElement == null)
            {
                return false;
            }

            var relId = (string?)sheetElement.Attribute(relIdAttributeName);
            if (string.IsNullOrWhiteSpace(relId))
            {
                throw new InvalidOperationException($"Sheet '{sheetName}' không có liên kết tới nội dung.");
            }

            var relationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), relId, StringComparison.Ordinal));

            var target = (string?)relationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                throw new InvalidOperationException($"Không xác định được đường dẫn nội dung cho sheet '{sheetName}'.");
            }

            var sheetPath = target.StartsWith("/", StringComparison.Ordinal)
                ? "xl" + target
                : "xl/" + target;

            sheetEntry = archive.GetEntry(sheetPath)
                ?? throw new InvalidOperationException($"File Excel template không hợp lệ: thiếu {sheetPath}.");

            using (var sheetReadStream = sheetEntry.Open())
            {
                sheetDoc = XDocument.Load(sheetReadStream);
            }

            sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                ?? throw new InvalidOperationException($"Sheet '{sheetName}' không chứa phần tử sheetData.");

            return true;
        }

        private static void RemoveRowsFrom(XElement sheetDataElement, XNamespace mainNs, int startRowIndex)
        {
            var rowsToRemove = sheetDataElement
                .Elements(mainNs + "row")
                .Where(r =>
                {
                    if (!int.TryParse((string?)r.Attribute("r"), out var rowIndex))
                    {
                        return false;
                    }

                    return rowIndex >= startRowIndex;
                })
                .ToList();

            foreach (var row in rowsToRemove)
            {
                row.Remove();
            }
        }

        private static string? BuildSummaryTitle(string? periodLabel)
        {
            if (!TryParsePeriod(periodLabel, out var month, out var year))
            {
                return null;
            }

            return $"BẢNG TỔNG HỢP HỘ TIÊU THỤ ĐIỆN THÁNG {month} NĂM {year}";
        }

        private static bool TryParsePeriod(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return false;
            }

            var match = Regex.Match(periodLabel, @"(\d{1,2})\s*/\s*(\d{4})");
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month) &&
                int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return month is >= 1 and <= 12 && year >= 2000;
            }

            match = Regex.Match(periodLabel, @"tháng\s*(\d{1,2}).*?(\d{4})", RegexOptions.IgnoreCase);
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month) &&
                int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return month is >= 1 and <= 12 && year >= 2000;
            }

            return false;
        }

        private static string NormalizeSheetKey(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static string? BuildSchoolPeriodText(string? periodLabel)
        {
            if (!TryParsePeriod(periodLabel, out var month, out var year))
            {
                return null;
            }

            return $"Tháng {month} năm {year}";
        }

        private static string BuildHanoiDateLine(DateTime date)
        {
            return $"Hà Nội, ngày {date.Day} tháng {date.Month} năm {date.Year}";
        }

        private static void ClearCellValue(XElement sheetDataElement, XNamespace ns, string cellReference)
        {
            UpdateNumberCell(sheetDataElement, ns, cellReference, value: null);
        }

        private static void UpdateNumberCell(XElement sheetDataElement, XNamespace ns, string cellReference, decimal? value)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return;
            }

            var rowIndex = GetRowIndex(cellReference);
            if (rowIndex <= 0)
            {
                return;
            }

            var row = sheetDataElement
                .Elements(ns + "row")
                .FirstOrDefault(r => string.Equals(
                    (string?)r.Attribute("r"),
                    rowIndex.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal));

            if (row == null)
            {
                return;
            }

            var cell = row
                .Elements(ns + "c")
                .FirstOrDefault(c => string.Equals(
                    (string?)c.Attribute("r"),
                    cellReference,
                    StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                cell = new XElement(ns + "c", new XAttribute("r", cellReference));
                row.Add(cell);
            }

            var styleAttr = (string?)cell.Attribute("s");

            cell.Attribute("t")?.Remove();
            cell.Elements(ns + "f").Remove();
            cell.Elements(ns + "v").Remove();
            cell.Elements(ns + "is").Remove();

            if (value == null)
            {
                if (!string.IsNullOrEmpty(styleAttr))
                {
                    cell.SetAttributeValue("s", styleAttr);
                }

                return;
            }

            cell.Add(new XElement(ns + "v", value.Value.ToString(CultureInfo.InvariantCulture)));

            if (!string.IsNullOrEmpty(styleAttr))
            {
                cell.SetAttributeValue("s", styleAttr);
            }
        }

        private static void UpdateTextCell(XElement sheetDataElement, XNamespace ns, string cellReference, string text)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return;
            }

            var rowIndex = GetRowIndex(cellReference);
            if (rowIndex <= 0)
            {
                return;
            }

            var row = sheetDataElement
                .Elements(ns + "row")
                .FirstOrDefault(r => string.Equals(
                    (string?)r.Attribute("r"),
                    rowIndex.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal));

            if (row == null)
            {
                return;
            }

            var cell = row
                .Elements(ns + "c")
                .FirstOrDefault(c => string.Equals(
                    (string?)c.Attribute("r"),
                    cellReference,
                    StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                cell = new XElement(ns + "c", new XAttribute("r", cellReference));
                row.Add(cell);
            }

            var styleAttr = (string?)cell.Attribute("s");

            cell.Attribute("t")?.Remove();
            cell.Elements(ns + "f").Remove();
            cell.Elements(ns + "v").Remove();
            cell.Elements(ns + "is").Remove();

            if (string.IsNullOrEmpty(text))
            {
                if (!string.IsNullOrEmpty(styleAttr))
                {
                    cell.SetAttributeValue("s", styleAttr);
                }

                return;
            }

            cell.SetAttributeValue("t", "inlineStr");
            cell.Add(new XElement(ns + "is", new XElement(ns + "t", text)));

            if (!string.IsNullOrEmpty(styleAttr))
            {
                cell.SetAttributeValue("s", styleAttr);
            }
        }

        private static int GetRowIndex(string cellReference)
        {
            var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex) ? rowIndex : 0;
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

        private static XElement CreateNumberCell(XNamespace ns, string column, int rowIndex, decimal? value)
        {
            var cellReference = column + rowIndex.ToString(CultureInfo.InvariantCulture);

            var cell = new XElement(ns + "c",
                new XAttribute("r", cellReference));

            if (value != null)
            {
                cell.Add(new XElement(ns + "v", value.Value.ToString(CultureInfo.InvariantCulture)));
            }

            return cell;
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
