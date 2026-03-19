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
    public static class ExcelRoundTripExportService
    {
        public sealed record ExportResult(
            int UpdatedRows,
            int UpdatedCells,
            int MissingCustomers,
            IReadOnlyList<int> MissingSequenceNumbers);

        public static ExportResult ExportLegacySummaryDataSheet(
            string templatePath,
            string outputPath,
            IReadOnlyList<Customer> customers,
            string? periodLabel)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
            {
                throw new ArgumentException("Template path is required.", nameof(templatePath));
            }

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("Template Excel file not found.", templatePath);
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("Output path is required.", nameof(outputPath));
            }

            if (customers == null)
            {
                throw new ArgumentNullException(nameof(customers));
            }

            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: thiếu xl/workbook.xml.");

            var workbookRelsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: thiếu xl/_rels/workbook.xml.rels.");

            var sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml")
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: thiếu xl/sharedStrings.xml.");

            XDocument workbookDoc;
            using (var stream = workbookEntry.Open())
            {
                workbookDoc = XDocument.Load(stream);
            }

            XDocument workbookRelsDoc;
            using (var stream = workbookRelsEntry.Open())
            {
                workbookRelsDoc = XDocument.Load(stream);
            }

            XDocument sharedStringsDoc;
            using (var stream = sharedStringsEntry.Open())
            {
                sharedStringsDoc = XDocument.Load(stream);
            }

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var sharedStringsRoot = sharedStringsDoc.Root
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: sharedStrings.xml bị rỗng.");

            var sharedStrings = LoadSharedStrings(sharedStringsRoot, mainNs);
            var sharedStringIndexByText = sharedStrings
                .Select((text, index) => (text, index))
                .GroupBy(x => x.text, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.First().index, StringComparer.Ordinal);

            if (!TryFindWorksheetPath(workbookDoc, workbookRelsDoc, mainNs, relNs, relPackageNs, sheetName: "Data", out var dataSheetPath))
            {
                throw new InvalidOperationException("File Excel template không có sheet 'Data'.");
            }

            var sheetEntry = archive.GetEntry(dataSheetPath)
                ?? throw new InvalidOperationException($"File Excel template không hợp lệ: thiếu {dataSheetPath}.");

            XDocument sheetDoc;
            using (var stream = sheetEntry.Open())
            {
                sheetDoc = XDocument.Load(stream);
            }

            var worksheetRoot = sheetDoc.Root
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: worksheet rỗng.");

            var sheetDataElement = worksheetRoot.Element(mainNs + "sheetData")
                ?? throw new InvalidOperationException("File Excel template không hợp lệ: thiếu sheetData.");

            var updatedCells = 0;

            // A1: Update title month/year.
            var desiredTitle = BuildLegacySummaryTitle(GetCellText(sheetDataElement, mainNs, "A1", sharedStrings), periodLabel);
            UpdateTextIfChanged(sheetDataElement, mainNs, "A1", desiredTitle, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);

            var rowIndexBySequence = BuildRowIndexBySequenceNumber(sheetDataElement, mainNs, startRowIndex: 5, sharedStrings);
            var missing = new List<int>();

            var updatedRows = 0;

            foreach (var customer in customers)
            {
                if (customer == null || customer.SequenceNumber <= 0)
                {
                    continue;
                }

                if (!rowIndexBySequence.TryGetValue(customer.SequenceNumber, out var rowIndex))
                {
                    missing.Add(customer.SequenceNumber);
                    continue;
                }

                var rowUpdated = false;

                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"B{rowIndex}", customer.Name, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"C{rowIndex}", customer.GroupName, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"D{rowIndex}", customer.Address, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"E{rowIndex}", customer.HouseholdPhone, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"F{rowIndex}", customer.RepresentativeName, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"G{rowIndex}", customer.Phone, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"H{rowIndex}", customer.BuildingName, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"J{rowIndex}", customer.MeterNumber, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"K{rowIndex}", customer.Category, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"L{rowIndex}", customer.Location, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"M{rowIndex}", customer.Substation, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"N{rowIndex}", customer.Page, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);
                rowUpdated |= UpdateTextIfChanged(sheetDataElement, mainNs, $"W{rowIndex}", customer.PerformedBy, sharedStrings, sharedStringsRoot, sharedStringIndexByText, ref updatedCells);

                rowUpdated |= UpdateNumberIfChanged(sheetDataElement, mainNs, $"O{rowIndex}", customer.CurrentIndex, ref updatedCells);
                rowUpdated |= UpdateNumberIfChanged(sheetDataElement, mainNs, $"P{rowIndex}", customer.PreviousIndex, ref updatedCells);
                rowUpdated |= UpdateNumberIfChanged(sheetDataElement, mainNs, $"Q{rowIndex}", customer.Multiplier, ref updatedCells);
                rowUpdated |= UpdateNumberIfChanged(sheetDataElement, mainNs, $"S{rowIndex}", customer.SubsidizedKwh, ref updatedCells);
                rowUpdated |= UpdateNumberIfChanged(sheetDataElement, mainNs, $"U{rowIndex}", customer.UnitPrice, ref updatedCells);

                if (rowUpdated)
                {
                    updatedRows++;
                }
            }

            using (var stream = sheetEntry.Open())
            {
                stream.SetLength(0);
                sheetDoc.Save(stream);
            }

            using (var stream = sharedStringsEntry.Open())
            {
                stream.SetLength(0);
                sharedStringsDoc.Save(stream);
            }

            return new ExportResult(
                UpdatedRows: updatedRows,
                UpdatedCells: updatedCells,
                MissingCustomers: missing.Count,
                MissingSequenceNumbers: missing);
        }

        private static bool TryFindWorksheetPath(
            XDocument workbookDoc,
            XDocument workbookRelsDoc,
            XNamespace mainNs,
            XNamespace relNs,
            XNamespace relPackageNs,
            string sheetName,
            out string sheetPath)
        {
            sheetPath = string.Empty;

            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets");
            if (sheetsElement == null)
            {
                return false;
            }

            var relIdAttrName = XName.Get("id", relNs.NamespaceName);

            var sheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheetElement == null)
            {
                return false;
            }

            var relId = (string?)sheetElement.Attribute(relIdAttrName);
            if (string.IsNullOrWhiteSpace(relId))
            {
                return false;
            }

            var relationshipsRoot = workbookRelsDoc.Root;
            if (relationshipsRoot == null)
            {
                return false;
            }

            var relationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), relId, StringComparison.Ordinal));

            var target = (string?)relationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                return false;
            }

            sheetPath = target.StartsWith("/", StringComparison.Ordinal)
                ? "xl" + target
                : "xl/" + target;

            return true;
        }

        private static List<string> LoadSharedStrings(XElement root, XNamespace mainNs)
        {
            return root
                .Elements(mainNs + "si")
                .Select(si =>
                    (string?)si.Element(mainNs + "t") ??
                    (string?)si.Element(mainNs + "r")?.Element(mainNs + "t") ??
                    si.Value ??
                    string.Empty)
                .ToList();
        }

        private static string? GetCellText(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            IReadOnlyList<string> sharedStrings)
        {
            var cell = GetCell(sheetDataElement, mainNs, cellReference, createIfMissing: false);
            if (cell == null)
            {
                return null;
            }

            var cellType = (string?)cell.Attribute("t");
            if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                var inlineText = cell.Element(mainNs + "is")?.Element(mainNs + "t")?.Value;
                return inlineText ?? string.Empty;
            }

            var v = cell.Element(mainNs + "v")?.Value;
            if (string.IsNullOrWhiteSpace(v))
            {
                return string.Empty;
            }

            if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedIndex) &&
                sharedIndex >= 0 &&
                sharedIndex < sharedStrings.Count)
            {
                return sharedStrings[sharedIndex] ?? string.Empty;
            }

            return v;
        }

        private static decimal? GetCellDecimal(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference)
        {
            var cell = GetCell(sheetDataElement, mainNs, cellReference, createIfMissing: false);
            if (cell == null)
            {
                return null;
            }

            if (cell.Element(mainNs + "f") != null)
            {
                return null;
            }

            var v = cell.Element(mainNs + "v")?.Value;
            if (string.IsNullOrWhiteSpace(v))
            {
                return null;
            }

            return decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
                ? value
                : null;
        }

        private static Dictionary<int, int> BuildRowIndexBySequenceNumber(
            XElement sheetDataElement,
            XNamespace mainNs,
            int startRowIndex,
            IReadOnlyList<string> sharedStrings)
        {
            var result = new Dictionary<int, int>();

            foreach (var row in sheetDataElement.Elements(mainNs + "row"))
            {
                var rowIndexText = (string?)row.Attribute("r");
                if (!int.TryParse(rowIndexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex) ||
                    rowIndex < startRowIndex)
                {
                    continue;
                }

                var cellRef = $"A{rowIndex}";
                var cellText = GetCellText(sheetDataElement, mainNs, cellRef, sharedStrings);
                if (!int.TryParse((cellText ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var sequence) ||
                    sequence <= 0)
                {
                    continue;
                }

                if (!result.ContainsKey(sequence))
                {
                    result[sequence] = rowIndex;
                }
            }

            return result;
        }

        private static bool UpdateTextIfChanged(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            string? newValue,
            List<string> sharedStrings,
            XElement sharedStringsRoot,
            Dictionary<string, int> sharedStringIndexByText,
            ref int updatedCells)
        {
            var current = GetCellText(sheetDataElement, mainNs, cellReference, sharedStrings) ?? string.Empty;
            var desired = (newValue ?? string.Empty).Trim();

            if (string.Equals(current?.Trim() ?? string.Empty, desired, StringComparison.Ordinal))
            {
                return false;
            }

            if (SetCellText(sheetDataElement, mainNs, cellReference, desired, sharedStringsRoot, sharedStrings, sharedStringIndexByText))
            {
                updatedCells++;
                return true;
            }

            return false;
        }

        private static bool UpdateNumberIfChanged(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            decimal? newValue,
            ref int updatedCells)
        {
            var current = GetCellDecimal(sheetDataElement, mainNs, cellReference);

            if (current == null && newValue == null)
            {
                return false;
            }

            if (current != null && newValue != null && current.Value == newValue.Value)
            {
                return false;
            }

            if (SetCellNumber(sheetDataElement, mainNs, cellReference, newValue))
            {
                updatedCells++;
                return true;
            }

            return false;
        }

        private static bool SetCellNumber(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            decimal? value)
        {
            var cell = GetCell(sheetDataElement, mainNs, cellReference, createIfMissing: true);
            if (cell == null)
            {
                return false;
            }

            if (cell.Element(mainNs + "f") != null)
            {
                return false;
            }

            cell.Attribute("t")?.Remove();
            cell.Elements(mainNs + "is").Remove();

            var vElement = cell.Element(mainNs + "v");
            if (value == null)
            {
                vElement?.Remove();
                return true;
            }

            var text = value.Value.ToString(CultureInfo.InvariantCulture);
            if (vElement == null)
            {
                cell.Add(new XElement(mainNs + "v", text));
            }
            else
            {
                vElement.Value = text;
            }

            return true;
        }

        private static bool SetCellText(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            string text,
            XElement sharedStringsRoot,
            List<string> sharedStrings,
            Dictionary<string, int> sharedStringIndexByText)
        {
            var cell = GetCell(sheetDataElement, mainNs, cellReference, createIfMissing: true);
            if (cell == null)
            {
                return false;
            }

            if (cell.Element(mainNs + "f") != null)
            {
                return false;
            }

            if (!sharedStringIndexByText.TryGetValue(text, out var index))
            {
                index = sharedStrings.Count;
                sharedStrings.Add(text);
                sharedStringIndexByText[text] = index;

                sharedStringsRoot.Add(new XElement(mainNs + "si", new XElement(mainNs + "t", text)));
                UpdateSharedStringCounts(sharedStringsRoot, sharedStrings.Count);
            }

            cell.SetAttributeValue("t", "s");
            cell.Elements(mainNs + "is").Remove();

            var vElement = cell.Element(mainNs + "v");
            if (vElement == null)
            {
                cell.Add(new XElement(mainNs + "v", index.ToString(CultureInfo.InvariantCulture)));
            }
            else
            {
                vElement.Value = index.ToString(CultureInfo.InvariantCulture);
            }

            return true;
        }

        private static void UpdateSharedStringCounts(XElement sharedStringsRoot, int uniqueCount)
        {
            sharedStringsRoot.SetAttributeValue("uniqueCount", uniqueCount.ToString(CultureInfo.InvariantCulture));

            var currentCount = (int?)sharedStringsRoot.Attribute("count");
            if (currentCount != null && currentCount.Value >= uniqueCount)
            {
                sharedStringsRoot.SetAttributeValue("count", currentCount.Value.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                sharedStringsRoot.SetAttributeValue("count", uniqueCount.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static XElement? GetCell(
            XElement sheetDataElement,
            XNamespace mainNs,
            string cellReference,
            bool createIfMissing)
        {
            var rowIndex = GetRowIndex(cellReference);
            if (rowIndex <= 0)
            {
                return null;
            }

            var row = sheetDataElement
                .Elements(mainNs + "row")
                .FirstOrDefault(r => string.Equals(
                    (string?)r.Attribute("r"),
                    rowIndex.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal));

            if (row == null)
            {
                return null;
            }

            var cell = row
                .Elements(mainNs + "c")
                .FirstOrDefault(c => string.Equals(
                    (string?)c.Attribute("r"),
                    cellReference,
                    StringComparison.OrdinalIgnoreCase));

            if (cell == null && createIfMissing)
            {
                cell = new XElement(mainNs + "c", new XAttribute("r", cellReference));
                row.Add(cell);
            }

            return cell;
        }

        private static int GetRowIndex(string cellReference)
        {
            var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex)
                ? rowIndex
                : 0;
        }

        private static string BuildLegacySummaryTitle(string? existingTitle, string? periodLabel)
        {
            var title = (existingTitle ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                return string.Empty;
            }

            if (!TryParsePeriodLabel(periodLabel, out var month, out var year))
            {
                return title;
            }

            var desired = $"THÁNG {month} NĂM {year}";
            var replaced = Regex.Replace(
                title,
                @"THÁNG\s+\d{1,2}\s+N[ĂA]M\s+\d{4}",
                desired,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            return string.Equals(replaced, title, StringComparison.Ordinal) ? title : replaced;
        }

        private static bool TryParsePeriodLabel(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            var text = (periodLabel ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var match = Regex.Match(text, @"(\d{1,2}).*?(\d{4})", RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            if (!int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month))
            {
                return false;
            }

            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return false;
            }

            return month is >= 1 and <= 12 && year is >= 1900 and <= 2200;
        }
    }
}
