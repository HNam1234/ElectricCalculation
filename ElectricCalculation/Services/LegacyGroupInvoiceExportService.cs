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
    public static class LegacyGroupInvoiceExportService
    {
        private static readonly string[] PreferredTemplateSheets =
        {
            "Trường Điện - Điện tử",
            "Trường Cơ khí"
        };

        private const int DetailStartRow = 13;
        private const int MaxDetailRows = 11; // legacy template uses rows 13..23

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

            var templateSheetBytes = ReadAllBytes(templateSheetEntry);

            var templateSheetIndex = TryParseWorksheetIndex(templateSheetPath);
            var templateSheetRelsBytes = templateSheetIndex > 0
                ? TryReadEntryBytes(archive, $"xl/worksheets/_rels/sheet{templateSheetIndex}.xml.rels")
                : null;

            var usedSheetIndexes = archive.Entries
                .Select(e => TryParseWorksheetIndex(e.FullName))
                .Where(i => i > 0)
                .ToHashSet();

            var maxRelId = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .Select(r => (string?)r.Attribute("Id"))
                .Where(id => id != null && Regex.IsMatch(id, @"^rId\\d+$"))
                .Select(id => int.Parse(id![3..], CultureInfo.InvariantCulture))
                .DefaultIfEmpty(0)
                .Max();

            var worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .Where(r =>
                    string.Equals((string?)r.Attribute("Type"), worksheetRelType, StringComparison.Ordinal) &&
                    !string.Equals((string?)r.Attribute("Id"), templateRelId, StringComparison.Ordinal))
                .Remove();

            sheetsElement.RemoveNodes();

            var chunks = list
                .Chunk(MaxDetailRows)
                .ToList();

            var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var groupSheetCount = 0;

            for (var chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
            {
                var sheetName = chunkIndex == 0
                    ? MakeUniqueSheetName(usedSheetNames, group)
                    : MakeUniqueSheetName(usedSheetNames, $"{group} ({chunkIndex + 1})");

                ZipArchiveEntry sheetEntry;
                var relId = templateRelId;

                if (chunkIndex == 0)
                {
                    sheetEntry = templateSheetEntry;
                }
                else
                {
                    var sheetIndex = 1;
                    while (usedSheetIndexes.Contains(sheetIndex))
                    {
                        sheetIndex++;
                    }

                    usedSheetIndexes.Add(sheetIndex);

                    var sheetPath = $"xl/worksheets/sheet{sheetIndex}.xml";
                    sheetEntry = archive.CreateEntry(sheetPath);
                    WriteAllBytes(sheetEntry, templateSheetBytes);

                    if (templateSheetRelsBytes != null)
                    {
                        var relsEntry = archive.CreateEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels");
                        WriteAllBytes(relsEntry, templateSheetRelsBytes);
                    }

                    relId = $"rId{++maxRelId}";
                    relationshipsRoot.Add(new XElement(relPackageNs + "Relationship",
                        new XAttribute("Id", relId),
                        new XAttribute("Type", worksheetRelType),
                        new XAttribute("Target", $"worksheets/sheet{sheetIndex}.xml")));

                    EnsureWorksheetContentType(contentTypesDoc, ctNs, sheetPath);
                }

                sheetsElement.Add(new XElement(mainNs + "sheet",
                    new XAttribute("name", sheetName),
                    new XAttribute("sheetId", (++groupSheetCount).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute(relIdAttributeName, relId)));

                XDocument sheetDoc;
                using (var sheetReadStream = sheetEntry.Open())
                {
                    sheetDoc = XDocument.Load(sheetReadStream);
                }

                var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                    ?? throw new InvalidOperationException("Legacy template worksheet has no sheetData section.");

                PopulateLegacySchoolLikeSheet(
                    sheetDataElement,
                    mainNs,
                    group,
                    chunks[chunkIndex],
                    periodLabel,
                    issuerName);

                using (var sheetWriteStream = sheetEntry.Open())
                {
                    sheetWriteStream.SetLength(0);
                    sheetDoc.Save(sheetWriteStream);
                }
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

            return chunks.Count;
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
            XElement sheetDataElement,
            XNamespace mainNs,
            string groupName,
            IReadOnlyList<Customer> customers,
            string periodLabel,
            string issuerName)
        {
            UpdateTextCell(sheetDataElement, mainNs, "J4", string.Empty);
            UpdateNumberCell(sheetDataElement, mainNs, "J6", null);

            var periodText = FormatPeriodLabel(periodLabel);
            if (!string.IsNullOrWhiteSpace(periodText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
            }

            var sharedAddress = GetSharedNonEmptyValue(customers, c => c.Address) ?? string.Empty;
            var sharedRepresentative = GetSharedNonEmptyValue(customers, c => c.RepresentativeName) ?? string.Empty;
            var sharedHouseholdPhone = GetSharedNonEmptyValue(customers, c => string.IsNullOrWhiteSpace(c.HouseholdPhone) ? c.Phone : c.HouseholdPhone) ?? string.Empty;
            var sharedRepresentativePhone = GetSharedNonEmptyValue(customers, c => c.Phone) ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(sharedHouseholdPhone) &&
                string.Equals(sharedHouseholdPhone, sharedRepresentativePhone, StringComparison.OrdinalIgnoreCase))
            {
                sharedRepresentativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(sharedHouseholdPhone) && !string.IsNullOrWhiteSpace(sharedRepresentativePhone))
            {
                sharedHouseholdPhone = sharedRepresentativePhone;
                sharedRepresentativePhone = string.Empty;
            }

            var representativeDisplay = !string.IsNullOrWhiteSpace(sharedRepresentative) ? sharedRepresentative : groupName;

            UpdateTextCell(sheetDataElement, mainNs, "A5", $"Kính gửi: {groupName}");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A7",
                string.IsNullOrWhiteSpace(sharedAddress) ? string.Empty : $"Địa chỉ hộ tiêu thụ: {sharedAddress}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A8",
                string.IsNullOrWhiteSpace(representativeDisplay) ? string.Empty : $"Đại diện: {representativeDisplay}.");
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

            for (var i = 0; i < MaxDetailRows; i++)
            {
                var rowIndex = DetailStartRow + i;

                if (i < customers.Count)
                {
                    var customer = customers[i];
                    UpdateNumberCell(sheetDataElement, mainNs, $"A{rowIndex}", i + 1);
                    UpdateNumberCell(sheetDataElement, mainNs, $"B{rowIndex}", customer.CurrentIndex);
                    UpdateNumberCell(sheetDataElement, mainNs, $"C{rowIndex}", customer.PreviousIndex);
                    UpdateNumberCell(sheetDataElement, mainNs, $"D{rowIndex}", customer.Multiplier <= 0 ? 1 : customer.Multiplier);
                    UpdateNumberCell(sheetDataElement, mainNs, $"F{rowIndex}", customer.SubsidizedKwh);
                    UpdateNumberCell(sheetDataElement, mainNs, $"G{rowIndex}", customer.UnitPrice);
                    UpdateTextCell(sheetDataElement, mainNs, $"I{rowIndex}", displayNames[i]);
                    UpdateTextCell(sheetDataElement, mainNs, $"J{rowIndex}", customer.Address ?? string.Empty);
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

            var totalAmount = customers.Sum(c => c.Amount);
            var amountText = VietnameseNumberTextService.ConvertAmountToText(totalAmount);
            UpdateTextCell(sheetDataElement, mainNs, "A25", string.IsNullOrWhiteSpace(amountText) ? string.Empty : $"Bằng chữ: {amountText}./.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "H27",
                $"Hà Nội, ngày {DateTime.Now.Day} tháng {DateTime.Now.Month} năm {DateTime.Now.Year}");

            var issuer = issuerName?.Trim() ?? string.Empty;
            UpdateTextCell(sheetDataElement, mainNs, "H32", issuer);
        }

        private static string FormatPeriodLabel(string periodLabel)
        {
            var text = (periodLabel ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(text) ? string.Empty : text;
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
            AddMultiHouseholdDetailPart(parts, ResolveMultiHouseholdLocation(customer), baseName);
            AddMultiHouseholdDetailPart(parts, customer.MeterNumber, baseName, "Công tơ ");
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
            var name = baseName;
            var suffix = 2;

            while (usedNames.Contains(name))
            {
                var tail = $" ({suffix})";
                var maxLen = Math.Max(1, 31 - tail.Length);
                var head = baseName.Length > maxLen ? baseName.Substring(0, maxLen).Trim() : baseName;
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

        private static int TryParseWorksheetIndex(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return 0;
            }

            var match = Regex.Match(path, @"sheet(\d+)\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return 0;
            }

            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) ? index : 0;
        }

        private static byte[]? TryReadEntryBytes(ZipArchive archive, string path)
        {
            var entry = archive.GetEntry(path);
            return entry == null ? null : ReadAllBytes(entry);
        }

        private static byte[] ReadAllBytes(ZipArchiveEntry entry)
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        private static void WriteAllBytes(ZipArchiveEntry entry, byte[] bytes)
        {
            using var stream = entry.Open();
            stream.SetLength(0);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void EnsureWorksheetContentType(XDocument contentTypesDoc, XNamespace ctNs, string worksheetPath)
        {
            var root = contentTypesDoc.Root
                ?? throw new InvalidOperationException("Legacy template is invalid: [Content_Types].xml root missing.");

            var partName = worksheetPath.StartsWith("/", StringComparison.Ordinal)
                ? worksheetPath
                : "/" + worksheetPath;

            var existing = root
                .Elements(ctNs + "Override")
                .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                return;
            }

            root.Add(new XElement(ctNs + "Override",
                new XAttribute("PartName", partName),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
        }
    }
}
