using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class LegacyGroupInvoiceExportService
    {
        private static readonly string[] PreferredTemplateSheets =
        {
            "Trường Điện - Điện tử",
            "Trường Cơ khí"
        };

        private const int DetailStartRow = 13;
        private const int TemplateDetailRowCount = 11; // legacy template uses rows 13..23

        public static int ExportGroupInvoice(
            string templatePath,
            string outputPath,
            string groupName,
            IReadOnlyList<Customer> customers,
            string periodLabel,
            string issuerName)
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

            var group = (groupName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(group))
            {
                group = "(Không có nhóm)";
            }

            var list = (customers ?? Array.Empty<Customer>())
                .Where(c => c != null)
                .OrderBy(c => c.SequenceNumber > 0 ? c.SequenceNumber : int.MaxValue)
                .ThenBy(c => c.Name)
                .ToList();

            if (list.Count == 0)
            {
                throw new ArgumentException("Customers list is empty.", nameof(customers));
            }

            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("Legacy template is invalid: missing xl/workbook.xml.");

            var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                ?? throw new InvalidOperationException("Legacy template is invalid: missing xl/_rels/workbook.xml.rels.");

            var contentTypesEntry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidOperationException("Legacy template is invalid: missing [Content_Types].xml.");

            XDocument workbookDoc;
            using (var workbookStream = workbookEntry.Open())
            {
                workbookDoc = XDocument.Load(workbookStream);
            }

            XDocument relDoc;
            using (var relStream = relEntry.Open())
            {
                relDoc = XDocument.Load(relStream);
            }

            XDocument contentTypesDoc;
            using (var ctStream = contentTypesEntry.Open())
            {
                contentTypesDoc = XDocument.Load(ctStream);
            }

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            XNamespace ctNs = "http://schemas.openxmlformats.org/package/2006/content-types";

            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets")
                ?? throw new InvalidOperationException("Legacy template is invalid: sheets collection not found.");

            var relationshipsRoot = relDoc.Root
                ?? throw new InvalidOperationException("Legacy template is invalid: relationships root missing.");

            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);

            var templateSheetElement = FindTemplateSheet(sheetsElement, mainNs);
            var templateRelId = (string?)templateSheetElement.Attribute(relIdAttributeName);
            if (string.IsNullOrWhiteSpace(templateRelId))
            {
                throw new InvalidOperationException("Legacy template is invalid: template sheet has no relationship id.");
            }

            var templateRelationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), templateRelId, StringComparison.Ordinal));

            var templateTarget = (string?)templateRelationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(templateTarget))
            {
                throw new InvalidOperationException("Legacy template is invalid: cannot locate template worksheet content.");
            }

            var templateSheetPath = templateTarget.StartsWith("/", StringComparison.Ordinal)
                ? "xl" + templateTarget
                : "xl/" + templateTarget;

            var templateSheetEntry = archive.GetEntry(templateSheetPath)
                ?? throw new InvalidOperationException($"Legacy template is invalid: missing {templateSheetPath}.");

            var worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .Where(r =>
                    string.Equals((string?)r.Attribute("Type"), worksheetRelType, StringComparison.Ordinal) &&
                    !string.Equals((string?)r.Attribute("Id"), templateRelId, StringComparison.Ordinal))
                .Remove();

            sheetsElement.RemoveNodes();
            var sheetName = MakeUniqueSheetName(new HashSet<string>(StringComparer.OrdinalIgnoreCase), group);
            sheetsElement.Add(new XElement(mainNs + "sheet",
                new XAttribute("name", sheetName),
                new XAttribute("sheetId", "1"),
                new XAttribute(relIdAttributeName, templateRelId)));

            RemoveDefinedNames(workbookDoc, mainNs);
            RemoveCalcChainArtifacts(archive, relationshipsRoot, contentTypesDoc, relPackageNs, ctNs);

            XDocument sheetDoc;
            using (var sheetReadStream = templateSheetEntry.Open())
            {
                sheetDoc = XDocument.Load(sheetReadStream);
            }

            var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                ?? throw new InvalidOperationException("Legacy template worksheet has no sheetData section.");

            PopulateLegacySchoolLikeSheet(
                sheetDoc,
                sheetDataElement,
                mainNs,
                group,
                list,
                periodLabel,
                issuerName);

            using (var sheetWriteStream = templateSheetEntry.Open())
            {
                sheetWriteStream.SetLength(0);
                sheetDoc.Save(sheetWriteStream);
            }

            using (var workbookWriteStream = workbookEntry.Open())
            {
                workbookWriteStream.SetLength(0);
                workbookDoc.Save(workbookWriteStream);
            }

            using (var relWriteStream = relEntry.Open())
            {
                relWriteStream.SetLength(0);
                relDoc.Save(relWriteStream);
            }

            using (var ctWriteStream = contentTypesEntry.Open())
            {
                ctWriteStream.SetLength(0);
                contentTypesDoc.Save(ctWriteStream);
            }

            return 1;
        }

        private static XElement FindTemplateSheet(XElement sheetsElement, XNamespace mainNs)
        {
            foreach (var name in PreferredTemplateSheets)
            {
                var found = sheetsElement
                    .Elements(mainNs + "sheet")
                    .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), name, StringComparison.OrdinalIgnoreCase));

                if (found != null)
                {
                    return found;
                }
            }

            var any = sheetsElement.Elements(mainNs + "sheet").FirstOrDefault();
            if (any != null)
            {
                return any;
            }

            throw new InvalidOperationException("Legacy template workbook has no sheets.");
        }

        private static void PopulateLegacySchoolLikeSheet(
            XDocument sheetDoc,
            XElement sheetDataElement,
            XNamespace mainNs,
            string groupName,
            IReadOnlyList<Customer> customers,
            string periodLabel,
            string issuerName)
        {
            var worksheetRoot = sheetDoc.Root
                ?? throw new InvalidOperationException("Legacy template worksheet is empty.");

            // Ensure numeric columns are wide enough so Excel doesn't display #### for large numbers.
            EnsureColumnMinWidth(worksheetRoot, sheetDataElement, mainNs, columnIndex: 2, minWidth: 11.5); // Chỉ số mới
            EnsureColumnMinWidth(worksheetRoot, sheetDataElement, mainNs, columnIndex: 3, minWidth: 11.5); // Chỉ số cũ
            EnsureColumnMinWidth(worksheetRoot, sheetDataElement, mainNs, columnIndex: 5, minWidth: 12); // Điện năng tiêu thụ
            EnsureColumnMinWidth(worksheetRoot, sheetDataElement, mainNs, columnIndex: 8, minWidth: 18); // Thành tiền

            var targetDetailRowCount = Math.Max(customers.Count, TemplateDetailRowCount);
            var extraDetailRows = targetDetailRowCount - TemplateDetailRowCount;
            if (extraDetailRows > 0)
            {
                ExpandDetailArea(
                    worksheetRoot,
                    sheetDataElement,
                    mainNs,
                    DetailStartRow + TemplateDetailRowCount,
                    extraDetailRows);
            }

            var totalRowIndex = DetailStartRow + targetDetailRowCount;
            var amountTextRowIndex = totalRowIndex + 1;
            var dateRowIndex = totalRowIndex + 3;
            var issuerRowIndex = totalRowIndex + 8;

            UpdateTextCell(sheetDataElement, mainNs, "J4", string.Empty);
            UpdateNumberCell(sheetDataElement, mainNs, "J6", null);

            var periodText = FormatPeriodLabel(periodLabel);
            if (!string.IsNullOrWhiteSpace(periodText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
            }

            var sharedAddress = ResolveGroupHeaderAddress(groupName, customers);
            var sharedRepresentative = GetSharedNonEmptyValue(customers, c => c.RepresentativeName) ?? string.Empty;
            var sharedHouseholdPhone = ResolveBestPhone(customers, c => c.HouseholdPhone);
            var sharedRepresentativePhone = ResolveBestPhone(customers, c => c.Phone);

            var normalizedHouseholdPhone = NormalizePhoneForComparison(sharedHouseholdPhone);
            var normalizedRepresentativePhone = NormalizePhoneForComparison(sharedRepresentativePhone);

            if (!string.IsNullOrWhiteSpace(normalizedHouseholdPhone) &&
                !string.IsNullOrWhiteSpace(normalizedRepresentativePhone) &&
                string.Equals(normalizedHouseholdPhone, normalizedRepresentativePhone, StringComparison.OrdinalIgnoreCase))
            {
                sharedRepresentativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(sharedHouseholdPhone) && !string.IsNullOrWhiteSpace(sharedRepresentativePhone))
            {
                sharedHouseholdPhone = sharedRepresentativePhone;
                sharedRepresentativePhone = string.Empty;
                normalizedHouseholdPhone = NormalizePhoneForComparison(sharedHouseholdPhone);
                normalizedRepresentativePhone = string.Empty;
            }

            var representativeDisplay = !string.IsNullOrWhiteSpace(sharedRepresentative) ? sharedRepresentative : groupName;

            UpdateTextCell(sheetDataElement, mainNs, "A5", $"Kính gửi: {groupName}");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A7",
                string.IsNullOrWhiteSpace(sharedAddress) ? string.Empty : EnsureTrailingPeriod($"Địa chỉ hộ tiêu thụ: {sharedAddress}"));
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A8",
                string.IsNullOrWhiteSpace(representativeDisplay) ? string.Empty : EnsureTrailingPeriod($"Đại diện: {representativeDisplay}"));
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "J7",
                string.IsNullOrWhiteSpace(sharedHouseholdPhone) ? string.Empty : $"Điện thoại: {sharedHouseholdPhone}");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "J8",
                string.IsNullOrWhiteSpace(sharedRepresentativePhone) ? string.Empty : $"Điện thoại: {sharedRepresentativePhone}");

            var displayNames = BuildMultiHouseholdDisplayNames(customers);

            for (var i = 0; i < targetDetailRowCount; i++)
            {
                var rowIndex = DetailStartRow + i;

                if (i < customers.Count)
                {
                    var customer = customers[i];
                    var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
                    UpdateNumberCell(sheetDataElement, mainNs, $"A{rowIndex}", i + 1);
                    UpdateNumberCell(sheetDataElement, mainNs, $"B{rowIndex}", customer.CurrentIndex);
                    UpdateNumberCell(sheetDataElement, mainNs, $"C{rowIndex}", customer.PreviousIndex);
                    UpdateNumberCell(sheetDataElement, mainNs, $"D{rowIndex}", multiplier);
                    UpdateNumberCell(sheetDataElement, mainNs, $"E{rowIndex}", customer.Consumption);
                    UpdateNumberCell(sheetDataElement, mainNs, $"F{rowIndex}", customer.SubsidizedKwh);
                    UpdateNumberCell(sheetDataElement, mainNs, $"G{rowIndex}", customer.UnitPrice);
                    UpdateNumberCell(sheetDataElement, mainNs, $"H{rowIndex}", customer.Amount);
                    UpdateTextCell(sheetDataElement, mainNs, $"I{rowIndex}", displayNames[i]);
                    UpdateTextCell(sheetDataElement, mainNs, $"J{rowIndex}", customer.Address ?? string.Empty);
                    continue;
                }

                ClearCellValue(sheetDataElement, mainNs, $"A{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"B{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"C{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"D{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"E{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"F{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"G{rowIndex}");
                ClearCellValue(sheetDataElement, mainNs, $"H{rowIndex}");
                UpdateTextCell(sheetDataElement, mainNs, $"I{rowIndex}", string.Empty);
                UpdateTextCell(sheetDataElement, mainNs, $"J{rowIndex}", string.Empty);
            }

            var totalAmount = customers.Sum(c => c.Amount);
            UpdateNumberCell(sheetDataElement, mainNs, $"H{totalRowIndex}", totalAmount);

            var amountText = VietnameseNumberTextService.ConvertAmountToText(totalAmount);
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"A{amountTextRowIndex}",
                string.IsNullOrWhiteSpace(amountText) ? string.Empty : $"Bằng chữ: {amountText}./.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"H{dateRowIndex}",
                $"Hà Nội, ngày {DateTime.Now.Day} tháng {DateTime.Now.Month} năm {DateTime.Now.Year}");

            var issuer = issuerName?.Trim() ?? string.Empty;
            UpdateTextCell(sheetDataElement, mainNs, $"H{issuerRowIndex}", issuer);
        }

        private static void ExpandDetailArea(
            XElement worksheetRoot,
            XElement sheetDataElement,
            XNamespace mainNs,
            int insertStartRow,
            int rowCount)
        {
            if (rowCount <= 0)
            {
                return;
            }

            var templateRow = sheetDataElement
                .Elements(mainNs + "row")
                .FirstOrDefault(row => TryParseRowIndex((string?)row.Attribute("r")) == insertStartRow - 1);

            if (templateRow == null)
            {
                throw new InvalidOperationException("Legacy template is invalid: cannot find detail row template.");
            }

            ShiftRows(sheetDataElement, mainNs, insertStartRow, rowCount);
            InsertClonedDetailRows(sheetDataElement, mainNs, templateRow, insertStartRow, rowCount);
            ShiftMergedRanges(worksheetRoot, mainNs, insertStartRow, rowCount);
            ShiftDimension(worksheetRoot, mainNs, insertStartRow, rowCount);
        }

        private static void EnsureColumnMinWidth(
            XElement worksheetRoot,
            XElement sheetDataElement,
            XNamespace mainNs,
            int columnIndex,
            double minWidth)
        {
            if (columnIndex <= 0 || minWidth <= 0)
            {
                return;
            }

            var cols = worksheetRoot.Element(mainNs + "cols");
            if (cols == null)
            {
                cols = new XElement(mainNs + "cols");
                sheetDataElement.AddBeforeSelf(cols);
            }

            var colElements = cols.Elements(mainNs + "col").ToList();
            XElement? covering = null;
            var coveringMin = 0;
            var coveringMax = 0;

            foreach (var col in colElements)
            {
                if (!int.TryParse((string?)col.Attribute("min"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var min))
                {
                    continue;
                }

                if (!int.TryParse((string?)col.Attribute("max"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var max))
                {
                    max = min;
                }

                if (columnIndex < min || columnIndex > max)
                {
                    continue;
                }

                covering = col;
                coveringMin = min;
                coveringMax = max;
                break;
            }

            if (covering == null)
            {
                var newCol = new XElement(mainNs + "col",
                    new XAttribute("min", columnIndex.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("max", columnIndex.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("width", minWidth.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("customWidth", "1"));

                var insertBefore = colElements
                    .Select(col => new { Col = col, MinText = (string?)col.Attribute("min") })
                    .Select(item => new
                    {
                        item.Col,
                        Min = int.TryParse(item.MinText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : int.MaxValue
                    })
                    .OrderBy(item => item.Min)
                    .FirstOrDefault(item => item.Min > columnIndex)?.Col;

                if (insertBefore != null)
                {
                    insertBefore.AddBeforeSelf(newCol);
                }
                else
                {
                    cols.Add(newCol);
                }

                return;
            }

            var existingWidthText = (string?)covering.Attribute("width");
            var existingWidth = double.TryParse(existingWidthText, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedWidth)
                ? parsedWidth
                : 0;

            if (existingWidth >= minWidth)
            {
                return;
            }

            if (coveringMin == columnIndex && coveringMax == columnIndex)
            {
                covering.SetAttributeValue("width", minWidth.ToString(CultureInfo.InvariantCulture));
                covering.SetAttributeValue("customWidth", "1");
                return;
            }

            var updatedWidth = minWidth.ToString(CultureInfo.InvariantCulture);

            XElement? left = null;
            if (coveringMin < columnIndex)
            {
                left = new XElement(covering);
                left.SetAttributeValue("min", coveringMin.ToString(CultureInfo.InvariantCulture));
                left.SetAttributeValue("max", (columnIndex - 1).ToString(CultureInfo.InvariantCulture));
            }

            var target = new XElement(covering);
            target.SetAttributeValue("min", columnIndex.ToString(CultureInfo.InvariantCulture));
            target.SetAttributeValue("max", columnIndex.ToString(CultureInfo.InvariantCulture));
            target.SetAttributeValue("width", updatedWidth);
            target.SetAttributeValue("customWidth", "1");

            XElement? right = null;
            if (columnIndex < coveringMax)
            {
                right = new XElement(covering);
                right.SetAttributeValue("min", (columnIndex + 1).ToString(CultureInfo.InvariantCulture));
                right.SetAttributeValue("max", coveringMax.ToString(CultureInfo.InvariantCulture));
            }

            if (left != null)
            {
                covering.AddBeforeSelf(left);
            }

            covering.AddBeforeSelf(target);

            if (right != null)
            {
                covering.AddBeforeSelf(right);
            }

            covering.Remove();
        }

        private static void ShiftRows(
            XElement sheetDataElement,
            XNamespace mainNs,
            int startRow,
            int offset)
        {
            var rowsToShift = sheetDataElement
                .Elements(mainNs + "row")
                .Select(row => new { Row = row, Index = TryParseRowIndex((string?)row.Attribute("r")) })
                .Where(item => item.Index >= startRow)
                .OrderByDescending(item => item.Index)
                .ToList();

            foreach (var item in rowsToShift)
            {
                var shiftedRow = item.Index + offset;
                item.Row.SetAttributeValue("r", shiftedRow.ToString(CultureInfo.InvariantCulture));

                foreach (var cell in item.Row.Elements(mainNs + "c"))
                {
                    var cellRef = (string?)cell.Attribute("r");
                    if (string.IsNullOrWhiteSpace(cellRef))
                    {
                        continue;
                    }

                    var shiftedRef = ShiftCellReference(cellRef, startRow, offset);
                    if (!string.Equals(cellRef, shiftedRef, StringComparison.Ordinal))
                    {
                        cell.SetAttributeValue("r", shiftedRef);
                    }
                }
            }
        }

        private static void InsertClonedDetailRows(
            XElement sheetDataElement,
            XNamespace mainNs,
            XElement templateRow,
            int startRow,
            int rowCount)
        {
            var firstShiftedRow = sheetDataElement
                .Elements(mainNs + "row")
                .FirstOrDefault(row => TryParseRowIndex((string?)row.Attribute("r")) == startRow + rowCount);

            for (var i = 0; i < rowCount; i++)
            {
                var rowIndex = startRow + i;
                var newRow = new XElement(templateRow);
                newRow.SetAttributeValue("r", rowIndex.ToString(CultureInfo.InvariantCulture));

                foreach (var cell in newRow.Elements(mainNs + "c"))
                {
                    var cellRef = (string?)cell.Attribute("r");
                    if (!string.IsNullOrWhiteSpace(cellRef))
                    {
                        cell.SetAttributeValue("r", ReplaceCellReferenceRow(cellRef, rowIndex));
                    }

                    cell.Attribute("t")?.Remove();
                    cell.Elements(mainNs + "f").Remove();
                    cell.Elements(mainNs + "v").Remove();
                    cell.Elements(mainNs + "is").Remove();
                }

                if (firstShiftedRow != null)
                {
                    firstShiftedRow.AddBeforeSelf(newRow);
                    continue;
                }

                sheetDataElement.Add(newRow);
            }
        }

        private static void ShiftMergedRanges(
            XElement worksheetRoot,
            XNamespace mainNs,
            int startRow,
            int offset)
        {
            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells == null)
            {
                return;
            }

            foreach (var mergeCell in mergeCells.Elements(mainNs + "mergeCell"))
            {
                var reference = (string?)mergeCell.Attribute("ref");
                if (string.IsNullOrWhiteSpace(reference))
                {
                    continue;
                }

                mergeCell.SetAttributeValue("ref", ShiftRangeReference(reference, startRow, offset));
            }
        }

        private static void ShiftDimension(
            XElement worksheetRoot,
            XNamespace mainNs,
            int startRow,
            int offset)
        {
            var dimension = worksheetRoot.Element(mainNs + "dimension");
            if (dimension == null)
            {
                return;
            }

            var reference = (string?)dimension.Attribute("ref");
            if (string.IsNullOrWhiteSpace(reference))
            {
                return;
            }

            dimension.SetAttributeValue("ref", ShiftRangeReference(reference, startRow, offset));
        }

        private static string ShiftRangeReference(string reference, int startRow, int offset)
        {
            var parts = reference.Split(':');
            if (parts.Length == 1)
            {
                return ShiftCellReference(parts[0], startRow, offset);
            }

            if (parts.Length == 2)
            {
                return $"{ShiftCellReference(parts[0], startRow, offset)}:{ShiftCellReference(parts[1], startRow, offset)}";
            }

            return reference;
        }

        private static string ShiftCellReference(string reference, int startRow, int offset)
        {
            if (!TrySplitCellReference(reference, out var column, out var row))
            {
                return reference;
            }

            var shiftedRow = row >= startRow ? row + offset : row;
            return $"{column}{shiftedRow}";
        }

        private static string ReplaceCellReferenceRow(string reference, int newRow)
        {
            if (!TrySplitCellReference(reference, out var column, out _))
            {
                return reference;
            }

            return $"{column}{newRow}";
        }

        private static bool TrySplitCellReference(string reference, out string column, out int row)
        {
            column = string.Empty;
            row = 0;

            var match = Regex.Match(reference ?? string.Empty, @"^([A-Za-z]+)(\d+)$", RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            column = match.Groups[1].Value.ToUpperInvariant();
            return int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
        }

        private static int TryParseRowIndex(string? rowIndexText)
        {
            return int.TryParse(rowIndexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
                ? value
                : 0;
        }

        private static string FormatPeriodLabel(string periodLabel)
        {
            var text = (periodLabel ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(text) ? string.Empty : text;
        }

        private static string ResolveGroupHeaderAddress(string groupName, IReadOnlyList<Customer> customers)
        {
            var shared = GetSharedNonEmptyValue(customers, c => c.Address);
            if (!string.IsNullOrWhiteSpace(shared))
            {
                return shared;
            }

            return (groupName ?? string.Empty).Trim();
        }

        private static string ResolveBestPhone(IReadOnlyList<Customer> customers, Func<Customer, string?> selector)
        {
            if (customers.Count == 0)
            {
                return string.Empty;
            }

            var groups = new Dictionary<string, (int Count, string Display, int DisplayScore)>(StringComparer.OrdinalIgnoreCase);

            foreach (var customer in customers)
            {
                var raw = selector(customer)?.Trim();
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                var normalized = NormalizePhoneForComparison(raw);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                if (!groups.TryGetValue(normalized, out var existing))
                {
                    groups[normalized] = (1, raw, ScorePhoneDisplay(raw));
                    continue;
                }

                var bestDisplay = existing.Display;
                var bestScore = existing.DisplayScore;
                var score = ScorePhoneDisplay(raw);
                if (score > bestScore)
                {
                    bestDisplay = raw;
                    bestScore = score;
                }

                groups[normalized] = (existing.Count + 1, bestDisplay, bestScore);
            }

            if (groups.Count == 0)
            {
                return string.Empty;
            }

            if (groups.Count == 1)
            {
                return groups.Values.First().Display;
            }

            var best = groups
                .OrderByDescending(kvp => kvp.Value.Count)
                .ThenByDescending(kvp => kvp.Value.DisplayScore)
                .First();

            return best.Value.Display;
        }

        private static int ScorePhoneDisplay(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
            {
                return -1;
            }

            var text = phone.Trim();
            var digitCount = text.Count(char.IsDigit);
            var score = digitCount;

            if (text.StartsWith("0", StringComparison.Ordinal))
            {
                score += 1000;
            }

            if (text.StartsWith("+", StringComparison.Ordinal))
            {
                score += 900;
            }

            if (text.StartsWith("84", StringComparison.Ordinal))
            {
                score += 800;
            }

            return score;
        }

        private static string NormalizePhoneForComparison(string phone)
        {
            var raw = phone?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(raw.Length);
            foreach (var ch in raw)
            {
                if (char.IsDigit(ch))
                {
                    builder.Append(ch);
                    continue;
                }
            }

            var digits = builder.ToString();
            if (string.IsNullOrWhiteSpace(digits))
            {
                return string.Empty;
            }

            if (digits.StartsWith("84", StringComparison.Ordinal) && digits.Length >= 9)
            {
                digits = "0" + digits.Substring(2);
            }

            digits = digits.TrimStart('0');
            return digits;
        }

        private static string EnsureTrailingPeriod(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var trimmed = text.Trim();
            return trimmed.EndsWith(".", StringComparison.Ordinal) ? trimmed : $"{trimmed}.";
        }

        private static void RemoveDefinedNames(XDocument workbookDoc, XNamespace mainNs)
        {
            workbookDoc.Root?
                .Element(mainNs + "definedNames")?
                .Remove();
        }

        private static void RemoveCalcChainArtifacts(
            ZipArchive archive,
            XElement relationshipsRoot,
            XDocument contentTypesDoc,
            XNamespace relPackageNs,
            XNamespace ctNs)
        {
            const string calcChainRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

            relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .Where(relationship => string.Equals(
                    (string?)relationship.Attribute("Type"),
                    calcChainRelationshipType,
                    StringComparison.Ordinal))
                .Remove();

            archive.GetEntry("xl/calcChain.xml")?.Delete();

            contentTypesDoc.Root?
                .Elements(ctNs + "Override")
                .Where(overrideNode =>
                    string.Equals((string?)overrideNode.Attribute("PartName"), "/xl/calcChain.xml", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals((string?)overrideNode.Attribute("PartName"), "xl/calcChain.xml", StringComparison.OrdinalIgnoreCase))
                .Remove();
        }

        private static string? GetSharedNonEmptyValue(IReadOnlyList<Customer> customers, Func<Customer, string?> selector)
        {
            string? shared = null;

            foreach (var customer in customers)
            {
                var value = selector(customer)?.Trim();
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (shared == null)
                {
                    shared = value;
                    continue;
                }

                if (!string.Equals(shared, value, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
            }

            return shared;
        }

        private static IReadOnlyList<string> BuildMultiHouseholdDisplayNames(IReadOnlyList<Customer> customers)
        {
            if (customers.Count == 0)
            {
                return Array.Empty<string>();
            }

            var baseNames = customers
                .Select((customer, index) => string.IsNullOrWhiteSpace(customer.Name) ? $"Hộ {index + 1}" : customer.Name.Trim())
                .ToArray();

            var displayNames = (string[])baseNames.Clone();

            var duplicateGroups = Enumerable.Range(0, customers.Count)
                .GroupBy(index => NormalizeKey(baseNames[index]), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key) && group.Count() > 1);

            foreach (var group in duplicateGroups)
            {
                var indexes = group.ToList();
                var detailPartsByIndex = indexes.ToDictionary(
                    index => index,
                    index => GetMultiHouseholdDetailParts(customers[index], baseNames[index]));
                var suffixDepthByIndex = indexes.ToDictionary(index => index, _ => 0);

                while (true)
                {
                    var collisions = indexes
                        .GroupBy(index => NormalizeKey(displayNames[index]), StringComparer.OrdinalIgnoreCase)
                        .Where(collision => collision.Count() > 1)
                        .Select(collision => collision.ToList())
                        .ToList();

                    if (collisions.Count == 0)
                    {
                        break;
                    }

                    var progressed = false;
                    foreach (var collision in collisions)
                    {
                        foreach (var index in collision)
                        {
                            if (suffixDepthByIndex[index] >= detailPartsByIndex[index].Count)
                            {
                                continue;
                            }

                            suffixDepthByIndex[index]++;
                            progressed = true;
                        }
                    }

                    foreach (var index in indexes)
                    {
                        displayNames[index] = ComposeMultiHouseholdDisplayName(
                            baseNames[index],
                            detailPartsByIndex[index],
                            suffixDepthByIndex[index]);
                    }

                    if (progressed)
                    {
                        continue;
                    }

                    foreach (var collision in collisions)
                    {
                        foreach (var index in collision)
                        {
                            displayNames[index] = $"{displayNames[index]} - Dòng {index + 1}";
                        }
                    }

                    break;
                }
            }

            return displayNames;
        }

        private static List<string> GetMultiHouseholdDetailParts(Customer customer, string baseName)
        {
            var parts = new List<string>();
            AddMultiHouseholdDetailPart(parts, customer.MeterNumber, baseName, "Công tơ ");
            AddMultiHouseholdDetailPart(parts, ResolveMultiHouseholdLocation(customer), baseName);
            AddMultiHouseholdDetailPart(parts, customer.BuildingName, baseName, "Mã sổ ");
            AddMultiHouseholdDetailPart(parts, customer.Address, baseName);
            return parts;
        }

        private static string? ResolveMultiHouseholdLocation(Customer customer)
        {
            var explicitLocation = customer.Location?.Trim();
            var inferredLocation = InferLocationFromAddress(customer.Address);

            if (!string.IsNullOrWhiteSpace(inferredLocation))
            {
                if (string.IsNullOrWhiteSpace(explicitLocation))
                {
                    return inferredLocation;
                }

                if (IsBasementLocation(inferredLocation) && !IsBasementLocation(explicitLocation))
                {
                    return inferredLocation;
                }
            }

            return string.IsNullOrWhiteSpace(explicitLocation) ? null : explicitLocation;
        }

        private static bool IsBasementLocation(string value)
        {
            return Regex.IsMatch(
                value,
                @"\b(thầm|tầng\s*hầm|tang\s*ham|hầm|ham)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string? InferLocationFromAddress(string? address)
        {
            var text = address?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var first = text
                .Split(new[] { " - ", " – ", " — ", "-", "–", "—" }, 2, StringSplitOptions.None)[0]
                .Trim();

            first = first.Split(new[] { ',' }, 2)[0].Trim();
            if (string.IsNullOrWhiteSpace(first))
            {
                return null;
            }

            if (IsBasementLocation(first))
            {
                return "Tầng hầm";
            }

            if (Regex.IsMatch(first, @"\b(tầng|tang|sân|san|trệt|tret|gác|gac|mái|mai)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return first;
            }

            return null;
        }

        private static string ComposeMultiHouseholdDisplayName(
            string baseName,
            IReadOnlyList<string> detailParts,
            int suffixDepth)
        {
            if (suffixDepth <= 0 || detailParts.Count == 0)
            {
                return baseName;
            }

            return $"{baseName} - {string.Join(" - ", detailParts.Take(suffixDepth))}";
        }

        private static void AddMultiHouseholdDetailPart(
            ICollection<string> parts,
            string? value,
            string baseName,
            string prefix = "")
        {
            var text = value?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            var candidate = string.IsNullOrWhiteSpace(prefix) ? text : $"{prefix}{text}";
            if (string.Equals(candidate, baseName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (parts.Any(part => string.Equals(part, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            parts.Add(candidate);
        }

        private static string NormalizeKey(string? value)
        {
            return value?.Trim() ?? string.Empty;
        }

        private static string MakeUniqueSheetName(ISet<string> usedNames, string baseName)
        {
            var sanitized = (baseName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "Không có nhóm";
            }

            foreach (var invalid in new[] { '\\', '/', '?', '*', '[', ']', ':' })
            {
                sanitized = sanitized.Replace(invalid, '_');
            }

            if (sanitized.Length > 31)
            {
                sanitized = sanitized[..31].Trim();
            }

            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "Không có nhóm";
            }

            var name = sanitized;
            var suffix = 2;

            while (usedNames.Contains(name))
            {
                var tail = $" ({suffix})";
                var maxLen = Math.Max(1, 31 - tail.Length);
                var head = sanitized.Length > maxLen ? sanitized.Substring(0, maxLen).Trim() : sanitized;
                name = head + tail;
                suffix++;
            }

            usedNames.Add(name);
            return name;
        }

        private static void UpdateTextCell(XElement sheetDataElement, XNamespace ns, string cellReference, string text)
        {
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

        private static void UpdateNumberCell(XElement sheetDataElement, XNamespace ns, string cellReference, decimal? value)
        {
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

        private static void ClearCellValue(XElement sheetDataElement, XNamespace ns, string cellReference)
        {
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
                return;
            }

            cell.Elements(ns + "f").Remove();
            cell.Elements(ns + "v").Remove();
            cell.Elements(ns + "is").Remove();
            cell.Attribute("t")?.Remove();
        }

        private static int GetRowIndex(string cellReference)
        {
            var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex)
                ? rowIndex
                : 0;
        }
    }
}
