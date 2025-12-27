using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    // Simple importer for the summary template:
    // Prefer sheet "Data" (fallback to a matching sheet), key columns:
    // A: No., B: Customer, C: Group, D: Address, E: Household phone, F: Representative,
    // G: Rep phone, H: Building, J: Meter, K: Category, L: Location, M: Substation, N: Page,
    // O: Current, P: Previous, Q: Multiplier, S: Subsidy, U: Unit price, W: Performed by.
    public static class ExcelImportService
    {
        public enum ImportField
        {
            SequenceNumber,
            Name,
            GroupName,
            Category,
            Address,
            RepresentativeName,
            HouseholdPhone,
            Phone,
            BuildingName,
            MeterNumber,
            Location,
            Substation,
            Page,
            CurrentIndex,
            PreviousIndex,
            Multiplier,
            SubsidizedKwh,
            UnitPrice,
            PerformedBy
        }

        private sealed record SheetInfo(string Name, string Path);

        private sealed record FieldRule(ImportField Field, string[] Keywords, int Priority, bool Required);

        private sealed record CompiledFieldRule(ImportField Field, string[] Keywords, int Priority, bool Required);

        private const double HeaderFieldScoreThreshold = 0.55;

        private static readonly IReadOnlyList<FieldRule> FieldRules = new[]
        {
            new FieldRule(ImportField.SequenceNumber, new[] { "stt", "so thu tu", "so tt", "no", "no." }, Priority: 10, Required: false),
            new FieldRule(ImportField.Name, new[] { "ten khach", "khach hang", "ho ten", "ho tieu thu", "ho tieu thu dien", "ten", "name" }, Priority: 10, Required: true),
            new FieldRule(ImportField.MeterNumber, new[] { "so cong to", "cong to", "meter", "meter number" }, Priority: 10, Required: false),
            new FieldRule(ImportField.CurrentIndex, new[] { "chi so moi", "cs moi", "current", "end" }, Priority: 10, Required: false),
            new FieldRule(ImportField.PreviousIndex, new[] { "chi so cu", "cs cu", "previous", "start" }, Priority: 10, Required: false),
            new FieldRule(ImportField.UnitPrice, new[] { "don gia", "gia", "unit price", "price" }, Priority: 10, Required: false),
            new FieldRule(ImportField.Multiplier, new[] { "he so", "hs", "multiplier" }, Priority: 10, Required: false),

            new FieldRule(ImportField.GroupName, new[] { "nhom", "don vi", "group", "unit" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Address, new[] { "dia chi ho tieu thu dien", "dia chi ho tieu thu", "dia chi", "dc", "address" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Phone, new[] { "dien thoai", "so dt", "sdt", "dt", "phone" }, Priority: 0, Required: false),
            new FieldRule(ImportField.HouseholdPhone, new[] { "dt ho", "dien thoai ho", "so dt ho" }, Priority: 0, Required: false),
            new FieldRule(ImportField.RepresentativeName, new[] { "dai dien", "nguoi dai dien", "representative" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Location, new[] { "vi tri lap dat cong to", "vi tri dat cong to", "vi tri lap dat", "vi tri", "location" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Substation, new[] { "tram", "tba", "substation" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Page, new[] { "trang", "page" }, Priority: 0, Required: false),
            new FieldRule(ImportField.SubsidizedKwh, new[] { "bao cap", "tro cap", "mien giam", "subsidy" }, Priority: 0, Required: false),
            new FieldRule(ImportField.PerformedBy, new[] { "nguoi thuc hien", "nguoi ghi chi so", "nguoi ghi", "nguoi thu", "performed" }, Priority: 0, Required: false),
            new FieldRule(ImportField.BuildingName, new[] { "toa", "nha", "ma so", "building", "book" }, Priority: 0, Required: false),
            new FieldRule(ImportField.Category, new[] { "loai", "category" }, Priority: 0, Required: false)
        };

        private static readonly IReadOnlyList<CompiledFieldRule> CompiledFieldRules = FieldRules
            .Select(rule => new CompiledFieldRule(
                rule.Field,
                rule.Keywords
                    .Select(NormalizeHeader)
                    .Where(keyword => !string.IsNullOrWhiteSpace(keyword))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray(),
                rule.Priority,
                rule.Required))
            .ToArray();

        private static readonly string[] PreferredNormalizedSheetNames =
        {
            "data",
            "bang ke",
            "ban in so"
        };

        public static ImportPreviewResult BuildPreview(string filePath)
        {
            return BuildPreview(filePath, sheetName: null, headerRowIndexOverride: null);
        }

        public static ImportPreviewResult BuildPreview(string filePath, string? sheetName)
        {
            return BuildPreview(filePath, sheetName, headerRowIndexOverride: null);
        }

        public static ImportPreviewResult BuildPreview(string filePath, string? sheetName, int? headerRowIndexOverride)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new InvalidOperationException("Duong dan file Excel dang trong.");
            }

            var forcedHeaderRowIndex = headerRowIndexOverride.HasValue && headerRowIndexOverride.Value > 0
                ? headerRowIndexOverride.Value
                : (int?)null;

            using var archive = ZipFile.OpenRead(filePath);

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var workbookDoc = LoadWorkbookDocument(archive);
            var workbookRelsDoc = LoadWorkbookRelationshipsDocument(archive);
            var sheets = LoadSheetInfos(workbookDoc, workbookRelsDoc, mainNs, relNs, relPackageNs);

            var selectedSheet = SelectSheet(sheets, sheetName, out var sheetWarningMessage);
            var sharedStrings = LoadSharedStrings(archive, mainNs);
            var sheetDataElement = LoadSheetDataElement(archive, mainNs, selectedSheet.Path);

            var firstPreviewRows = new List<ImportSampleRow>(capacity: 10);
            var headerPreviewRows = new List<ImportSampleRow>(capacity: 10);
            var sampleRows = new List<ImportSampleRow>(capacity: 5);
            var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Dictionary<string, string?>? headerCells = null;
            int? headerRowIndex = null;
            int? dataStartRowIndex = null;
            Dictionary<ImportField, string>? detectedHeaderMap = null;

            foreach (var row in sheetDataElement.Elements(mainNs + "row"))
            {
                if (!int.TryParse((string?)row.Attribute("r"), out var rowIndex))
                {
                    continue;
                }

                var cells = ReadRowCells(row, mainNs, sharedStrings);
                foreach (var key in cells.Keys)
                {
                    columns.Add(key);
                }

                if (firstPreviewRows.Count < 10)
                {
                    firstPreviewRows.Add(new ImportSampleRow(rowIndex, cells));
                }

                if (forcedHeaderRowIndex.HasValue && rowIndex == forcedHeaderRowIndex.Value)
                {
                    headerRowIndex = rowIndex;
                    dataStartRowIndex = rowIndex + 1;
                    headerCells = new Dictionary<string, string?>(cells, StringComparer.OrdinalIgnoreCase);

                    if (TryDetectHeaderMap(cells, out var forcedMap))
                    {
                        detectedHeaderMap = forcedMap;
                    }
                }

                if (forcedHeaderRowIndex == null &&
                    headerRowIndex == null &&
                    rowIndex <= 200 &&
                    TryDetectHeaderMap(cells, out var detectedMap))
                {
                    headerRowIndex = rowIndex;
                    dataStartRowIndex = rowIndex + 1;
                    detectedHeaderMap = detectedMap;
                    headerCells = new Dictionary<string, string?>(cells, StringComparer.OrdinalIgnoreCase);
                }

                if (headerRowIndex != null &&
                    rowIndex >= headerRowIndex.Value &&
                    headerPreviewRows.Count < 10)
                {
                    headerPreviewRows.Add(new ImportSampleRow(rowIndex, cells));
                }

                if (dataStartRowIndex != null &&
                    rowIndex >= dataStartRowIndex.Value &&
                    sampleRows.Count < 5)
                {
                    sampleRows.Add(new ImportSampleRow(rowIndex, cells));
                }

                if (forcedHeaderRowIndex != null)
                {
                    var reachedPreview = firstPreviewRows.Count >= 10;
                    var reachedHeader = headerRowIndex != null;
                    var reachedForcedHeaderPreview = headerPreviewRows.Count >= 10;
                    var reachedSamples = sampleRows.Count >= 5;

                    if (reachedPreview && reachedHeader && reachedForcedHeaderPreview && reachedSamples)
                    {
                        break;
                    }

                    continue;
                }

                var reachedTopPreview = firstPreviewRows.Count >= 10;
                var reachedScan = rowIndex > 200 || headerRowIndex != null;
                var reachedHeaderPreview = headerRowIndex == null || headerPreviewRows.Count >= 10;
                var reachedSamplesAuto = headerRowIndex == null || sampleRows.Count >= 5;
                if (reachedTopPreview && reachedScan && reachedHeaderPreview && reachedSamplesAuto)
                {
                    break;
                }
            }

            var previewRows = headerPreviewRows.Count > 0 ? headerPreviewRows : firstPreviewRows;

            if (forcedHeaderRowIndex != null && headerRowIndex == null)
            {
                throw new InvalidOperationException($"Khong tim thay dong tieu de so {forcedHeaderRowIndex.Value} trong sheet.");
            }

            if (headerCells == null && previewRows.Count > 0)
            {
                headerCells = new Dictionary<string, string?>(previewRows[0].Cells, StringComparer.OrdinalIgnoreCase);
            }

            if (sampleRows.Count == 0 && previewRows.Count > 1)
            {
                var assumedHeaderRow = headerRowIndex ?? previewRows[0].RowIndex;
                foreach (var row in previewRows.Where(r => r.RowIndex > assumedHeaderRow).Take(5))
                {
                    sampleRows.Add(row);
                }
            }

            var orderedColumns = columns
                .OrderBy(GetColumnIndex)
                .ToArray();

            var previewTable = BuildPreviewDataTable(orderedColumns, previewRows, headerCells);

            var columnPreviews = new List<ImportColumnPreview>(orderedColumns.Length);
            foreach (var column in orderedColumns)
            {
                headerCells ??= new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

                var headerText = headerCells.TryGetValue(column, out var headerValue) ? headerValue ?? string.Empty : string.Empty;
                var normalizedHeader = NormalizeHeader(headerText);

                ImportField? suggestedField = null;
                var suggestedScore = 0d;

                var scoredFields = ScoreFields(normalizedHeader);
                if (scoredFields.Count > 0 && scoredFields[0].score >= HeaderFieldScoreThreshold)
                {
                    suggestedField = scoredFields[0].field;
                    suggestedScore = scoredFields[0].score;
                }

                var sampleValues = GetSampleValues(sampleRows, column, maxCount: 5);
                columnPreviews.Add(new ImportColumnPreview(column, headerText, sampleValues, suggestedField, suggestedScore));
            }

            ApplyAmbiguousIndexHeuristics(columnPreviews);
            ApplyRepresentativePhoneHeuristics(columnPreviews);

            var headerSignature = headerRowIndex == null
                ? null
                : ComputeHeaderSignature(orderedColumns.Select(c =>
                    NormalizeHeader(headerCells != null && headerCells.TryGetValue(c, out var value) ? value : null)));

            var warningMessage = AppendWarning(sheetWarningMessage, detectedHeaderMap != null ? BuildHeaderMissingColumnsWarning(detectedHeaderMap) : null);

            return new ImportPreviewResult(
                filePath,
                sheets.Select(s => s.Name).ToArray(),
                selectedSheet.Name,
                headerRowIndex,
                dataStartRowIndex,
                headerSignature,
                columnPreviews,
                previewTable,
                sampleRows,
                warningMessage);
        }

        public static IList<Customer> ImportFromFile(
            string filePath,
            Dictionary<ImportField, string> confirmedMap,
            int? dataStartRowIndex,
            out string? warningMessage,
            out ImportRunReport report)
        {
            return ImportFromFile(filePath, sheetName: null, confirmedMap, dataStartRowIndex, out warningMessage, out report);
        }

        public static IList<Customer> ImportFromFile(
            string filePath,
            string? sheetName,
            Dictionary<ImportField, string> confirmedMap,
            int? dataStartRowIndex,
            out string? warningMessage,
            out ImportRunReport report)
        {
            warningMessage = null;
            var warnings = new List<string>();
            var errors = new List<string>();
            var result = new List<Customer>();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new InvalidOperationException("Duong dan file Excel dang trong.");
            }

            using var archive = ZipFile.OpenRead(filePath);

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var workbookDoc = LoadWorkbookDocument(archive);
            var workbookRelsDoc = LoadWorkbookRelationshipsDocument(archive);
            var sheets = LoadSheetInfos(workbookDoc, workbookRelsDoc, mainNs, relNs, relPackageNs);

            var selectedSheet = SelectSheet(sheets, sheetName, out var sheetWarningMessage);
            warningMessage = AppendWarning(warningMessage, sheetWarningMessage);

            var sharedStrings = LoadSharedStrings(archive, mainNs);
            var sheetDataElement = LoadSheetDataElement(archive, mainNs, selectedSheet.Path);

            var headerMap = confirmedMap ?? new Dictionary<ImportField, string>();
            var allowTemplateFallbackColumns = false;
            string Fallback(string column) => allowTemplateFallbackColumns ? column : string.Empty;
            var fallbackSequenceNumber = 1;

            var totalRows = 0;
            var importedRows = 0;
            var skippedRows = 0;

            foreach (var row in sheetDataElement.Elements(mainNs + "row"))
            {
                if (!int.TryParse((string?)row.Attribute("r"), out var rowIndex))
                {
                    continue;
                }

                var cells = ReadRowCells(row, mainNs, sharedStrings);

                if (dataStartRowIndex == null && rowIndex <= 200 && TryDetectHeaderMap(cells, out _))
                {
                    dataStartRowIndex = rowIndex + 1;
                    continue;
                }

                var sequenceNumber = GetMappedInt(cells, headerMap, ImportField.SequenceNumber, fallbackColumn: Fallback("A"));

                if (dataStartRowIndex == null)
                {
                    if (sequenceNumber > 0)
                    {
                        dataStartRowIndex = rowIndex;
                    }
                    else
                    {
                        continue;
                    }
                }

                if (rowIndex < dataStartRowIndex.Value)
                {
                    continue;
                }

                totalRows++;

                if (sequenceNumber <= 0)
                {
                    if (!headerMap.ContainsKey(ImportField.SequenceNumber))
                    {
                        var hasSomeData =
                            !string.IsNullOrWhiteSpace(GetMappedString(cells, headerMap, ImportField.Name, fallbackColumn: Fallback("B"))) ||
                            !string.IsNullOrWhiteSpace(GetMappedString(cells, headerMap, ImportField.MeterNumber, fallbackColumn: Fallback("J")));

                        if (!hasSomeData)
                        {
                            skippedRows++;
                            continue;
                        }

                        sequenceNumber = fallbackSequenceNumber++;
                    }
                    else
                    {
                        skippedRows++;
                        continue;
                    }
                }

                if (!headerMap.ContainsKey(ImportField.SequenceNumber))
                {
                    fallbackSequenceNumber = Math.Max(fallbackSequenceNumber, sequenceNumber + 1);
                }

                var householdPhone = (GetMappedString(cells, headerMap, ImportField.HouseholdPhone, fallbackColumn: Fallback("E")) ?? string.Empty).Trim();
                var representativePhone = (GetMappedString(cells, headerMap, ImportField.Phone, fallbackColumn: Fallback("G")) ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(representativePhone))
                {
                    representativePhone = householdPhone;
                }

                var representativeName = (GetMappedString(cells, headerMap, ImportField.RepresentativeName, fallbackColumn: Fallback("F")) ?? string.Empty).Trim();
                var householdName = (GetMappedString(cells, headerMap, ImportField.Name, fallbackColumn: Fallback("B")) ?? string.Empty).Trim();

                var customer = new Customer
                {
                    SequenceNumber = sequenceNumber,
                    Name = householdName,
                    GroupName = (GetMappedString(cells, headerMap, ImportField.GroupName, fallbackColumn: Fallback("C")) ?? string.Empty).Trim(),
                    Address = (GetMappedString(cells, headerMap, ImportField.Address, fallbackColumn: Fallback("D")) ?? string.Empty).Trim(),
                    RepresentativeName = representativeName,
                    HouseholdPhone = householdPhone,
                    Phone = representativePhone,
                    BuildingName = (GetMappedString(cells, headerMap, ImportField.BuildingName, fallbackColumn: Fallback("H")) ?? string.Empty).Trim(),
                    MeterNumber = (GetMappedString(cells, headerMap, ImportField.MeterNumber, fallbackColumn: Fallback("J")) ?? string.Empty).Trim(),
                    Category = (GetMappedString(cells, headerMap, ImportField.Category, fallbackColumn: Fallback("K")) ?? string.Empty).Trim(),
                    Location = (GetMappedString(cells, headerMap, ImportField.Location, fallbackColumn: Fallback("L")) ?? string.Empty).Trim()
                };

                customer.Substation = (GetMappedString(cells, headerMap, ImportField.Substation, fallbackColumn: Fallback("M")) ?? string.Empty).Trim();
                customer.Page = (GetMappedString(cells, headerMap, ImportField.Page, fallbackColumn: Fallback("N")) ?? string.Empty).Trim();

                customer.CurrentIndex = GetMappedNullableDecimal(cells, headerMap, ImportField.CurrentIndex, fallbackColumn: Fallback("O"));
                customer.PreviousIndex = GetMappedDecimal(cells, headerMap, ImportField.PreviousIndex, fallbackColumn: Fallback("P"));
                customer.Multiplier = GetMappedDecimal(cells, headerMap, ImportField.Multiplier, fallbackColumn: Fallback("Q"));
                customer.SubsidizedKwh = GetMappedDecimal(cells, headerMap, ImportField.SubsidizedKwh, fallbackColumn: Fallback("S"));
                customer.UnitPrice = GetMappedDecimal(cells, headerMap, ImportField.UnitPrice, fallbackColumn: Fallback("U"));

                customer.PerformedBy = (GetMappedString(cells, headerMap, ImportField.PerformedBy, fallbackColumn: Fallback("W")) ?? string.Empty).Trim();

                var unitPriceText = GetMappedString(cells, headerMap, ImportField.UnitPrice, fallbackColumn: Fallback("U"));
                if (!string.IsNullOrWhiteSpace(unitPriceText) &&
                    !decimal.TryParse(unitPriceText, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    warnings.Add($"Dong {rowIndex}: Don gia '{unitPriceText}' khong doc duoc (co the bi tinh = 0).");
                }

                if (customer.Multiplier <= 0)
                {
                    warnings.Add($"Dong {rowIndex}: He so <= 0, tu dong set = 1.");
                    customer.Multiplier = 1;
                }

                if (customer.CurrentIndex.HasValue && customer.CurrentIndex.Value < customer.PreviousIndex)
                {
                    warnings.Add($"Dong {rowIndex}: Chi so moi < chi so cu (moi={customer.CurrentIndex:0.##}, cu={customer.PreviousIndex:0.##}).");
                }

                result.Add(customer);
                importedRows++;
            }

            report = new ImportRunReport(
                TotalRows: totalRows,
                ImportedRows: importedRows,
                SkippedRows: skippedRows,
                WarningCount: warnings.Count,
                ErrorCount: errors.Count,
                Warnings: warnings,
                Errors: errors);

            return result;
        }

        private static XDocument LoadWorkbookDocument(ZipArchive archive)
        {
            var workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry == null)
            {
                throw new InvalidOperationException("File Excel khong hop le: thieu xl/workbook.xml.");
            }

            return XDocument.Load(workbookEntry.Open());
        }

        private static XDocument LoadWorkbookRelationshipsDocument(ZipArchive archive)
        {
            var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (relEntry == null)
            {
                throw new InvalidOperationException("File Excel khong hop le: thieu xl/_rels/workbook.xml.rels.");
            }

            return XDocument.Load(relEntry.Open());
        }

        private static IReadOnlyList<SheetInfo> LoadSheetInfos(
            XDocument workbookDoc,
            XDocument relDoc,
            XNamespace mainNs,
            XNamespace relNs,
            XNamespace relPackageNs)
        {
            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets");
            if (sheetsElement == null)
            {
                throw new InvalidOperationException("File Excel khong hop le: khong tim thay danh sach sheet.");
            }

            var relationships = relDoc
                .Root?
                .Elements(relPackageNs + "Relationship")
                .Select(r => new
                {
                    Id = (string?)r.Attribute("Id"),
                    Target = (string?)r.Attribute("Target")
                })
                .Where(r => !string.IsNullOrWhiteSpace(r.Id) && !string.IsNullOrWhiteSpace(r.Target))
                .ToDictionary(r => r.Id!, r => r.Target!, StringComparer.Ordinal)
                ?? new Dictionary<string, string>(StringComparer.Ordinal);

            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);
            var result = new List<SheetInfo>();

            foreach (var sheet in sheetsElement.Elements(mainNs + "sheet"))
            {
                var name = (string?)sheet.Attribute("name") ?? string.Empty;
                var relId = (string?)sheet.Attribute(relIdAttributeName);
                if (string.IsNullOrWhiteSpace(relId))
                {
                    continue;
                }

                if (!relationships.TryGetValue(relId, out var target) || string.IsNullOrWhiteSpace(target))
                {
                    continue;
                }

                var sheetPath = target.StartsWith("/", StringComparison.Ordinal)
                    ? "xl" + target
                    : "xl/" + target;

                result.Add(new SheetInfo(name, sheetPath));
            }

            return result;
        }

        private static SheetInfo SelectSheet(IReadOnlyList<SheetInfo> sheets, string? requestedSheetName, out string? warningMessage)
        {
            warningMessage = null;
            if (sheets == null || sheets.Count == 0)
            {
                throw new InvalidOperationException("Khong tim thay sheet nao trong workbook.");
            }

            if (!string.IsNullOrWhiteSpace(requestedSheetName))
            {
                var match = sheets.FirstOrDefault(s => string.Equals(s.Name, requestedSheetName, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    return match;
                }

                warningMessage = $"Khong tim thay sheet '{requestedSheetName}'. Dang dung sheet '{sheets[0].Name}'.";
            }

            var dataSheet = sheets.FirstOrDefault(s => string.Equals(NormalizeHeader(s.Name), "data", StringComparison.OrdinalIgnoreCase));
            if (dataSheet != null)
            {
                return dataSheet;
            }

            var preferred = sheets.FirstOrDefault(s => PreferredNormalizedSheetNames.Contains(NormalizeHeader(s.Name), StringComparer.OrdinalIgnoreCase));
            if (preferred != null)
            {
                return preferred;
            }

            warningMessage ??= $"Khong tim thay sheet 'Data'. Dang doc tu sheet '{sheets[0].Name}'.";
            return sheets[0];
        }

        private static List<string> LoadSharedStrings(ZipArchive archive, XNamespace mainNs)
        {
            var sharedStrings = new List<string>();
            var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedEntry == null)
            {
                return sharedStrings;
            }

            var sharedDoc = XDocument.Load(sharedEntry.Open());
            foreach (var si in sharedDoc.Root!.Elements(mainNs + "si"))
            {
                var textParts = si
                    .Descendants(mainNs + "t")
                    .Select(t => (string?)t ?? string.Empty);

                sharedStrings.Add(string.Concat(textParts));
            }

            return sharedStrings;
        }

        private static XElement LoadSheetDataElement(ZipArchive archive, XNamespace mainNs, string sheetPath)
        {
            if (string.IsNullOrWhiteSpace(sheetPath))
            {
                throw new InvalidOperationException("Khong xac dinh duoc duong dan sheet trong file Excel.");
            }

            var sheetEntry = archive.GetEntry(sheetPath);
            if (sheetEntry == null)
            {
                throw new InvalidOperationException($"File Excel khong hop le: thieu {sheetPath}.");
            }

            var sheetDoc = XDocument.Load(sheetEntry.Open());
            var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData");
            if (sheetDataElement == null)
            {
                throw new InvalidOperationException($"Sheet '{sheetPath}' khong chua du lieu (sheetData).");
            }

            return sheetDataElement;
        }

        private static Dictionary<string, string?> ReadRowCells(XElement row, XNamespace mainNs, IReadOnlyList<string> sharedStrings)
        {
            var cells = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            foreach (var cell in row.Elements(mainNs + "c"))
            {
                var reference = (string?)cell.Attribute("r");
                if (string.IsNullOrWhiteSpace(reference))
                {
                    continue;
                }

                var columnLetters = new string(reference.TakeWhile(char.IsLetter).ToArray());
                if (string.IsNullOrEmpty(columnLetters))
                {
                    continue;
                }

                var type = (string?)cell.Attribute("t");
                string? cellValue = null;

                if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
                {
                    var textParts = cell
                        .Descendants(mainNs + "is")
                        .Descendants(mainNs + "t")
                        .Select(t => (string?)t ?? string.Empty);

                    cellValue = string.Concat(textParts);
                }
                else
                {
                    var valueElement = cell.Element(mainNs + "v");
                    if (valueElement != null)
                    {
                        var rawValue = valueElement.Value;

                        if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) &&
                                index >= 0 &&
                                index < sharedStrings.Count)
                            {
                                cellValue = sharedStrings[index];
                            }
                            else
                            {
                                cellValue = rawValue;
                            }
                        }
                        else
                        {
                            cellValue = rawValue;
                        }
                    }
                }

                cells[columnLetters] = cellValue;
            }

            return cells;
        }

        private static DataTable BuildPreviewDataTable(
            string[] orderedColumns,
            IReadOnlyList<ImportSampleRow> previewRows,
            IReadOnlyDictionary<string, string?>? headerCells)
        {
            var table = new DataTable();
            table.Columns.Add("Row", typeof(int)).Caption = "Dòng";

            foreach (var column in orderedColumns)
            {
                var dataColumn = table.Columns.Add(column, typeof(string));
                if (headerCells != null &&
                    headerCells.TryGetValue(column, out var headerText) &&
                    !string.IsNullOrWhiteSpace(headerText))
                {
                    dataColumn.Caption = $"{column} - {headerText.Trim()}";
                }
            }

            foreach (var row in previewRows)
            {
                var dataRow = table.NewRow();
                dataRow["Row"] = row.RowIndex;

                foreach (var column in orderedColumns)
                {
                    dataRow[column] = row.Cells.TryGetValue(column, out var value) ? value ?? string.Empty : string.Empty;
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static IReadOnlyList<string> GetSampleValues(IReadOnlyList<ImportSampleRow> sampleRows, string column, int maxCount)
        {
            if (maxCount <= 0 || sampleRows.Count == 0)
            {
                return Array.Empty<string>();
            }

            var values = new List<string>(capacity: maxCount);
            foreach (var row in sampleRows)
            {
                if (values.Count >= maxCount)
                {
                    break;
                }

                if (!row.Cells.TryGetValue(column, out var value) || string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                values.Add(value.Trim());
            }

            return values;
        }

        private static void ApplyAmbiguousIndexHeuristics(List<ImportColumnPreview> columnPreviews)
        {
            if (columnPreviews == null || columnPreviews.Count < 2)
            {
                return;
            }

            var hasStrongIndexSuggestion = columnPreviews.Any(c =>
                (c.SuggestedField == ImportField.CurrentIndex || c.SuggestedField == ImportField.PreviousIndex) &&
                c.SuggestedScore >= HeaderFieldScoreThreshold);

            if (hasStrongIndexSuggestion)
            {
                return;
            }

            var ambiguous = columnPreviews
                .Where(c => IsAmbiguousIndexHeader(c.HeaderText))
                .OrderBy(c => GetColumnIndex(c.ColumnLetter))
                .Take(2)
                .ToList();

            if (ambiguous.Count < 2)
            {
                return;
            }

            var left = ambiguous[0];
            var right = ambiguous[1];

            for (var i = 0; i < columnPreviews.Count; i++)
            {
                if (string.Equals(columnPreviews[i].ColumnLetter, left.ColumnLetter, StringComparison.OrdinalIgnoreCase))
                {
                    columnPreviews[i] = columnPreviews[i] with
                    {
                        SuggestedField = ImportField.PreviousIndex,
                        SuggestedScore = HeaderFieldScoreThreshold
                    };
                }

                if (string.Equals(columnPreviews[i].ColumnLetter, right.ColumnLetter, StringComparison.OrdinalIgnoreCase))
                {
                    columnPreviews[i] = columnPreviews[i] with
                    {
                        SuggestedField = ImportField.CurrentIndex,
                        SuggestedScore = HeaderFieldScoreThreshold
                    };
                }
            }
        }

        private static void ApplyRepresentativePhoneHeuristics(List<ImportColumnPreview> columnPreviews)
        {
            if (columnPreviews == null || columnPreviews.Count == 0)
            {
                return;
            }

            for (var i = 0; i < columnPreviews.Count; i++)
            {
                var header = NormalizeHeader(columnPreviews[i].HeaderText);
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                if (IsRepresentativePhoneHeader(header))
                {
                    columnPreviews[i] = columnPreviews[i] with
                    {
                        SuggestedField = ImportField.Phone,
                        SuggestedScore = Math.Max(columnPreviews[i].SuggestedScore, HeaderFieldScoreThreshold)
                    };
                    continue;
                }

                if (IsHouseholdPhoneHeader(header))
                {
                    columnPreviews[i] = columnPreviews[i] with
                    {
                        SuggestedField = ImportField.HouseholdPhone,
                        SuggestedScore = Math.Max(columnPreviews[i].SuggestedScore, HeaderFieldScoreThreshold)
                    };
                    continue;
                }

                if (IsRepresentativeNameHeader(header))
                {
                    columnPreviews[i] = columnPreviews[i] with
                    {
                        SuggestedField = ImportField.RepresentativeName,
                        SuggestedScore = Math.Max(columnPreviews[i].SuggestedScore, HeaderFieldScoreThreshold)
                    };
                }
            }
        }

        private static bool IsRepresentativePhoneHeader(string normalizedHeader)
        {
            if (string.IsNullOrWhiteSpace(normalizedHeader))
            {
                return false;
            }

            var hasPhone = normalizedHeader.Contains("dien thoai", StringComparison.OrdinalIgnoreCase) ||
                           normalizedHeader.Contains("sdt", StringComparison.OrdinalIgnoreCase);

            if (!hasPhone)
            {
                return false;
            }

            return normalizedHeader.Contains("nguoi dai dien", StringComparison.OrdinalIgnoreCase) ||
                   normalizedHeader.Contains("dai dien", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHouseholdPhoneHeader(string normalizedHeader)
        {
            if (string.IsNullOrWhiteSpace(normalizedHeader))
            {
                return false;
            }

            var hasPhone = normalizedHeader.Contains("dien thoai", StringComparison.OrdinalIgnoreCase) ||
                           normalizedHeader.Contains("sdt", StringComparison.OrdinalIgnoreCase);

            if (!hasPhone)
            {
                return false;
            }

            return normalizedHeader.Contains("ho tieu thu", StringComparison.OrdinalIgnoreCase) ||
                   normalizedHeader.Contains("ho gia dinh", StringComparison.OrdinalIgnoreCase) ||
                   normalizedHeader.Contains("dien thoai ho", StringComparison.OrdinalIgnoreCase) ||
                   normalizedHeader.Contains("dt ho", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRepresentativeNameHeader(string normalizedHeader)
        {
            if (string.IsNullOrWhiteSpace(normalizedHeader))
            {
                return false;
            }

            var hasName = normalizedHeader.Contains("dai dien", StringComparison.OrdinalIgnoreCase) ||
                          normalizedHeader.Contains("nguoi dai dien", StringComparison.OrdinalIgnoreCase);

            if (!hasName)
            {
                return false;
            }

            var hasPhone = normalizedHeader.Contains("dien thoai", StringComparison.OrdinalIgnoreCase) ||
                           normalizedHeader.Contains("sdt", StringComparison.OrdinalIgnoreCase);

            return !hasPhone;
        }

        private static bool IsAmbiguousIndexHeader(string headerText)
        {
            var normalized = NormalizeHeader(headerText);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var hasChiSo = normalized.Contains("chi so", StringComparison.OrdinalIgnoreCase) ||
                          normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries).Contains("cs", StringComparer.OrdinalIgnoreCase);

            if (!hasChiSo)
            {
                return false;
            }

            return !ContainsAnyToken(normalized, "moi", "cu", "current", "previous", "start", "end");
        }

        private static bool ContainsAnyToken(string text, params string[] tokens)
        {
            if (string.IsNullOrWhiteSpace(text) || tokens == null || tokens.Length == 0)
            {
                return false;
            }

            var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
                if (string.IsNullOrWhiteSpace(token))
                {
                    continue;
                }

                if (parts.Contains(token, StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string ComputeHeaderSignature(IEnumerable<string> normalizedHeaders)
        {
            var payload = string.Join("|", normalizedHeaders ?? Array.Empty<string>());
            var bytes = Encoding.UTF8.GetBytes(payload);
            var hash = SHA256.HashData(bytes);
            return Convert.ToHexString(hash);
        }

        public static IList<Customer> ImportFromFile(string filePath, out string? warningMessage)
        {
            warningMessage = null;
            var result = new List<Customer>();

            using var archive = ZipFile.OpenRead(filePath);

            var workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry == null)
            {
                throw new InvalidOperationException("File Excel không hợp lệ: thiếu xl/workbook.xml.");
            }

            var workbookDoc = XDocument.Load(workbookEntry.Open());
            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets");
            if (sheetsElement == null)
            {
                throw new InvalidOperationException("File Excel không hợp lệ: không tìm thấy danh sách sheet.");
            }

            var preferredSheetNames = new[]
            {
                "Data",
                "Bảng kê",
                "Bảng kê",
                "Bang ke",
                "Ban  in so",
                "Ban in so"
            };

            var dataSheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), "Data", StringComparison.OrdinalIgnoreCase))
                ?? sheetsElement
                    .Elements(mainNs + "sheet")
                    .FirstOrDefault(s =>
                    {
                        var name = (string?)s.Attribute("name");
                        return !string.IsNullOrWhiteSpace(name) &&
                               preferredSheetNames.Any(preferred => string.Equals(name, preferred, StringComparison.OrdinalIgnoreCase));
                    })
                ?? sheetsElement.Elements(mainNs + "sheet").FirstOrDefault();

            var selectedSheetName = (string?)dataSheetElement?.Attribute("name") ?? "Data";
            if (!string.Equals(selectedSheetName, "Data", StringComparison.OrdinalIgnoreCase) &&
                !preferredSheetNames.Any(preferred => string.Equals(selectedSheetName, preferred, StringComparison.OrdinalIgnoreCase)))
            {
                warningMessage = $"Không tìm thấy sheet 'Data'. Đang import từ sheet '{selectedSheetName}'.";
            }

            if (dataSheetElement == null)
            {
                throw new InvalidOperationException("Không tìm thấy sheet 'Data' trong file Excel.");
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
                throw new InvalidOperationException("File Excel không hợp lệ: thiếu xl/_rels/workbook.xml.rels.");
            }

            var relDoc = XDocument.Load(relEntry.Open());
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
                throw new InvalidOperationException($"File Excel không hợp lệ: thiếu {sheetPath}.");
            }

            // Shared strings
            var sharedStrings = new List<string>();
            var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedEntry != null)
            {
                var sharedDoc = XDocument.Load(sharedEntry.Open());
                foreach (var si in sharedDoc.Root!.Elements(mainNs + "si"))
                {
                    var textParts = si
                        .Descendants(mainNs + "t")
                        .Select(t => (string?)t ?? string.Empty);

                    sharedStrings.Add(string.Concat(textParts));
                }
            }

            var sheetDoc = XDocument.Load(sheetEntry.Open());
            var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData");
            if (sheetDataElement == null)
            {
                throw new InvalidOperationException("Sheet 'Data' không chứa dữ liệu (sheetData).");
            }

            int? dataStartRowIndex = null;
            Dictionary<ImportField, string>? headerMap = null;
            var fallbackSequenceNumber = 1;

            foreach (var row in sheetDataElement.Elements(mainNs + "row"))
            {
                if (!int.TryParse((string?)row.Attribute("r"), out var rowIndex))
                {
                    continue;
                }

                var cells = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

                foreach (var cell in row.Elements(mainNs + "c"))
                {
                    var reference = (string?)cell.Attribute("r");
                    if (string.IsNullOrWhiteSpace(reference))
                    {
                        continue;
                    }

                    var columnLetters = new string(reference.TakeWhile(char.IsLetter).ToArray());
                    if (string.IsNullOrEmpty(columnLetters))
                    {
                        continue;
                    }

                    var type = (string?)cell.Attribute("t");
                    string? cellValue = null;

                    if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
                    {
                        var textParts = cell
                            .Descendants(mainNs + "is")
                            .Descendants(mainNs + "t")
                            .Select(t => (string?)t ?? string.Empty);

                        cellValue = string.Concat(textParts);
                    }
                    else
                    {
                        var valueElement = cell.Element(mainNs + "v");
                        if (valueElement != null)
                        {
                            var rawValue = valueElement.Value;

                            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
                            {
                                if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) &&
                                    index >= 0 &&
                                    index < sharedStrings.Count)
                                {
                                    cellValue = sharedStrings[index];
                                }
                                else
                                {
                                    cellValue = rawValue;
                                }
                            }
                            else
                            {
                                cellValue = rawValue;
                            }
                        }
                    }

                    cells[columnLetters] = cellValue;
                }

                if (headerMap == null && rowIndex <= 200 && TryDetectHeaderMap(cells, out var detectedMap))
                {
                    headerMap = detectedMap;
                    dataStartRowIndex = rowIndex + 1;
                    warningMessage = AppendWarning(warningMessage, BuildHeaderMissingColumnsWarning(headerMap));
                    continue;
                }

                var sequenceNumber = GetMappedInt(cells, headerMap, ImportField.SequenceNumber, fallbackColumn: "A");

                // Detect the first data row dynamically (avoid hardcoding a template row index).
                // The first row with a positive "STT" is considered the start of data.
                if (dataStartRowIndex == null)
                {
                    if (sequenceNumber > 0)
                    {
                        dataStartRowIndex = rowIndex;
                    }
                    else
                    {
                        continue;
                    }
                }

                if (rowIndex < dataStartRowIndex.Value)
                {
                    continue;
                }

                // Skip rows without sequence number.
                if (sequenceNumber <= 0)
                {
                    if (headerMap != null && !headerMap.ContainsKey(ImportField.SequenceNumber))
                    {
                        var hasSomeData =
                            !string.IsNullOrWhiteSpace(GetMappedString(cells, headerMap, ImportField.Name, fallbackColumn: "B")) ||
                            !string.IsNullOrWhiteSpace(GetMappedString(cells, headerMap, ImportField.MeterNumber, fallbackColumn: "J"));

                        if (!hasSomeData)
                        {
                            continue;
                        }

                        sequenceNumber = fallbackSequenceNumber++;
                    }
                    else
                    {
                        continue;
                    }
                }

                if (headerMap != null && !headerMap.ContainsKey(ImportField.SequenceNumber))
                {
                    fallbackSequenceNumber = Math.Max(fallbackSequenceNumber, sequenceNumber + 1);
                }

                var householdPhone = (GetMappedString(cells, headerMap, ImportField.HouseholdPhone, fallbackColumn: "E") ?? string.Empty).Trim();
                var representativePhone = (GetMappedString(cells, headerMap, ImportField.Phone, fallbackColumn: "G") ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(representativePhone))
                {
                    representativePhone = householdPhone;
                }

                var representativeName = (GetMappedString(cells, headerMap, ImportField.RepresentativeName, fallbackColumn: "F") ?? string.Empty).Trim();
                var householdName = (GetMappedString(cells, headerMap, ImportField.Name, fallbackColumn: "B") ?? string.Empty).Trim();

                var customer = new Customer
                {
                    SequenceNumber = sequenceNumber,
                    Name = householdName,
                    GroupName = (GetMappedString(cells, headerMap, ImportField.GroupName, fallbackColumn: "C") ?? string.Empty).Trim(),
                    Address = (GetMappedString(cells, headerMap, ImportField.Address, fallbackColumn: "D") ?? string.Empty).Trim(),
                    RepresentativeName = representativeName,
                    HouseholdPhone = householdPhone,
                    Phone = representativePhone,
                    BuildingName = (GetMappedString(cells, headerMap, ImportField.BuildingName, fallbackColumn: "H") ?? string.Empty).Trim(),
                    MeterNumber = (GetMappedString(cells, headerMap, ImportField.MeterNumber, fallbackColumn: "J") ?? string.Empty).Trim(),
                    Category = (GetMappedString(cells, headerMap, ImportField.Category, fallbackColumn: "K") ?? string.Empty).Trim(),
                    Location = (GetMappedString(cells, headerMap, ImportField.Location, fallbackColumn: "L") ?? string.Empty).Trim()
                };

                customer.Substation = (GetMappedString(cells, headerMap, ImportField.Substation, fallbackColumn: "M") ?? string.Empty).Trim();
                customer.Page = (GetMappedString(cells, headerMap, ImportField.Page, fallbackColumn: "N") ?? string.Empty).Trim();

                customer.CurrentIndex = GetMappedNullableDecimal(cells, headerMap, ImportField.CurrentIndex, fallbackColumn: "O");
                customer.PreviousIndex = GetMappedDecimal(cells, headerMap, ImportField.PreviousIndex, fallbackColumn: "P");
                customer.Multiplier = GetMappedDecimal(cells, headerMap, ImportField.Multiplier, fallbackColumn: "Q");
                customer.SubsidizedKwh = GetMappedDecimal(cells, headerMap, ImportField.SubsidizedKwh, fallbackColumn: "S");
                customer.UnitPrice = GetMappedDecimal(cells, headerMap, ImportField.UnitPrice, fallbackColumn: "U");

                customer.PerformedBy = (GetMappedString(cells, headerMap, ImportField.PerformedBy, fallbackColumn: "W") ?? string.Empty).Trim();

                if (customer.Multiplier <= 0)
                {
                    customer.Multiplier = 1;
                }

                result.Add(customer);
            }

            if (headerMap == null)
            {
                warningMessage = AppendWarning(
                    warningMessage,
                    "Khong thay dong tieu de; app se doc theo cot mau (A,B,C...). Neu file khong theo template, hay them 1 dong header.");
            }

            return result;
        }

        private static string? AppendWarning(string? existing, string? message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return existing;
            }

            if (string.IsNullOrWhiteSpace(existing))
            {
                return message;
            }

            return $"{existing}\n{message}";
        }

        private static string? GetMappedString(
            IDictionary<string, string?> cells,
            IDictionary<ImportField, string>? map,
            ImportField field,
            string fallbackColumn)
        {
            if (map != null && map.TryGetValue(field, out var column) && !string.IsNullOrWhiteSpace(column))
            {
                return GetString(cells, column);
            }

            return string.IsNullOrWhiteSpace(fallbackColumn) ? null : GetString(cells, fallbackColumn);
        }

        private static int GetMappedInt(
            IDictionary<string, string?> cells,
            IDictionary<ImportField, string>? map,
            ImportField field,
            string fallbackColumn)
        {
            if (map != null && map.TryGetValue(field, out var column) && !string.IsNullOrWhiteSpace(column))
            {
                return GetInt(cells, column);
            }

            return string.IsNullOrWhiteSpace(fallbackColumn) ? 0 : GetInt(cells, fallbackColumn);
        }

        private static decimal GetMappedDecimal(
            IDictionary<string, string?> cells,
            IDictionary<ImportField, string>? map,
            ImportField field,
            string fallbackColumn)
        {
            if (map != null && map.TryGetValue(field, out var column) && !string.IsNullOrWhiteSpace(column))
            {
                return GetDecimal(cells, column);
            }

            return string.IsNullOrWhiteSpace(fallbackColumn) ? 0 : GetDecimal(cells, fallbackColumn);
        }

        private static decimal? GetMappedNullableDecimal(
            IDictionary<string, string?> cells,
            IDictionary<ImportField, string>? map,
            ImportField field,
            string fallbackColumn)
        {
            if (map != null && map.TryGetValue(field, out var column) && !string.IsNullOrWhiteSpace(column))
            {
                return GetNullableDecimal(cells, column);
            }

            return string.IsNullOrWhiteSpace(fallbackColumn) ? null : GetNullableDecimal(cells, fallbackColumn);
        }

        private static bool TryDetectHeaderMap(IDictionary<string, string?> cells, out Dictionary<ImportField, string> map)
        {
            map = new Dictionary<ImportField, string>();

            var bestByField = new Dictionary<ImportField, (string column, double score)>();

            foreach (var pair in cells.OrderBy(pair => GetColumnIndex(pair.Key)))
            {
                var normalized = NormalizeHeader(pair.Value);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                var scoredFields = ScoreFields(normalized);
                if (scoredFields.Count == 0)
                {
                    continue;
                }

                var best = scoredFields[0];
                if (best.score < HeaderFieldScoreThreshold)
                {
                    continue;
                }

                if (!bestByField.TryGetValue(best.field, out var current) ||
                    best.score > current.score ||
                    (Math.Abs(best.score - current.score) < 0.0001 &&
                     GetColumnIndex(pair.Key) < GetColumnIndex(current.column)))
                {
                    bestByField[best.field] = (pair.Key, best.score);
                }
            }

            foreach (var pair in bestByField)
            {
                map[pair.Key] = pair.Value.column;
            }

            var hasRequiredFields = true;
            foreach (var rule in CompiledFieldRules)
            {
                if (!rule.Required)
                {
                    continue;
                }

                if (!map.ContainsKey(rule.Field))
                {
                    hasRequiredFields = false;
                    break;
                }
            }

            var hasAnyKeyField =
                map.ContainsKey(ImportField.MeterNumber) ||
                map.ContainsKey(ImportField.CurrentIndex) ||
                map.ContainsKey(ImportField.PreviousIndex) ||
                map.ContainsKey(ImportField.UnitPrice);

            if (!hasRequiredFields || !hasAnyKeyField)
            {
                map.Clear();
                return false;
            }

            var isHeader = map.Count >= 4 || map.ContainsKey(ImportField.SequenceNumber);
            if (!isHeader)
            {
                map.Clear();
                return false;
            }

            return true;
        }

        private static string? BuildHeaderMissingColumnsWarning(IDictionary<ImportField, string> headerMap)
        {
            var importantFields = new Dictionary<ImportField, string>
            {
                [ImportField.UnitPrice] = "Don gia",
                [ImportField.Multiplier] = "He so",
                [ImportField.CurrentIndex] = "Chi so moi",
                [ImportField.PreviousIndex] = "Chi so cu"
            };

            var missingLabels = importantFields
                .Where(pair => !headerMap.ContainsKey(pair.Key))
                .Select(pair => pair.Value)
                .ToArray();

            if (missingLabels.Length == 0)
            {
                return null;
            }

            return $"Da nhan dien dong tieu de nhung thieu cot: {string.Join(", ", missingLabels)}.";
        }

        private static IReadOnlyList<(ImportField field, double score)> ScoreFields(string normalizedHeader)
        {
            if (string.IsNullOrWhiteSpace(normalizedHeader))
            {
                return Array.Empty<(ImportField field, double score)>();
            }

            var header = normalizedHeader.Trim();
            if (header.Length == 0)
            {
                return Array.Empty<(ImportField field, double score)>();
            }

            var headerWords = header.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var headerWordSet = new HashSet<string>(headerWords, StringComparer.OrdinalIgnoreCase);

            var results = new List<(ImportField field, double score, int priority)>();

            foreach (var rule in CompiledFieldRules)
            {
                double rawScore = 0;

                foreach (var keyword in rule.Keywords)
                {
                    if (string.IsNullOrWhiteSpace(keyword))
                    {
                        continue;
                    }

                    var wordCount = CountWords(keyword);

                    if (IsFullKeywordMatch(header, headerWordSet, keyword))
                    {
                        rawScore += 1.0 + 0.15 * Math.Max(0, wordCount - 1);
                        continue;
                    }

                    if (IsPartialKeywordMatch(headerWords, keyword))
                    {
                        rawScore += 0.45 + 0.08 * Math.Max(0, wordCount - 1);
                    }
                }

                if (rawScore <= 0)
                {
                    continue;
                }

                var normalizedScore = rawScore / (rawScore + 0.5);
                var score = Math.Min(1.0, normalizedScore + rule.Priority * 0.005);

                results.Add((rule.Field, score, rule.Priority));
            }

            return results
                .OrderByDescending(pair => pair.score)
                .ThenByDescending(pair => pair.priority)
                .Select(pair => (pair.field, pair.score))
                .ToArray();
        }

        private static int GetColumnIndex(string columnLetters)
        {
            if (string.IsNullOrWhiteSpace(columnLetters))
            {
                return int.MaxValue;
            }

            var value = 0;
            foreach (var ch in columnLetters.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z')
                {
                    break;
                }

                value = value * 26 + (ch - 'A' + 1);
            }

            return value == 0 ? int.MaxValue : value;
        }

        private static int CountWords(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return 0;
            }

            return input.Count(ch => ch == ' ') + 1;
        }

        private static bool IsFullKeywordMatch(string normalizedHeader, HashSet<string> headerWords, string normalizedKeyword)
        {
            if (string.IsNullOrWhiteSpace(normalizedHeader) || string.IsNullOrWhiteSpace(normalizedKeyword))
            {
                return false;
            }

            if (normalizedKeyword.IndexOf(' ') < 0)
            {
                return headerWords.Contains(normalizedKeyword);
            }

            if (ContainsPhraseWithWordBoundary(normalizedHeader, normalizedKeyword))
            {
                return true;
            }

            var collapsed = normalizedKeyword.Replace(" ", string.Empty);
            return !string.IsNullOrWhiteSpace(collapsed) && headerWords.Contains(collapsed);
        }

        private static bool ContainsPhraseWithWordBoundary(string text, string phrase)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(phrase))
            {
                return false;
            }

            var index = text.IndexOf(phrase, StringComparison.OrdinalIgnoreCase);
            while (index >= 0)
            {
                var startOk = index == 0 || text[index - 1] == ' ';
                var endIndex = index + phrase.Length;
                var endOk = endIndex == text.Length || text[endIndex] == ' ';

                if (startOk && endOk)
                {
                    return true;
                }

                index = text.IndexOf(phrase, index + 1, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static bool IsPartialKeywordMatch(string[] headerWords, string normalizedKeyword)
        {
            if (headerWords.Length == 0 || string.IsNullOrWhiteSpace(normalizedKeyword))
            {
                return false;
            }

            foreach (var word in headerWords)
            {
                if (word.Length <= normalizedKeyword.Length)
                {
                    continue;
                }

                if (word.Contains(normalizedKeyword, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeHeader(string? input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var text = input.Trim().ToLowerInvariant();
            text = text.Replace('đ', 'd');

            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalized.Length);

            foreach (var ch in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category == UnicodeCategory.NonSpacingMark)
                {
                    continue;
                }

                if (char.IsLetterOrDigit(ch) || ch == '%' || ch == ' ')
                {
                    sb.Append(ch);
                }
                else
                {
                    sb.Append(' ');
                }
            }

            return string.Join(' ', sb.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries));
        }

        private static string? GetString(IDictionary<string, string?> cells, string column)
        {
            return cells.TryGetValue(column, out var value) ? value : null;
        }

        private static int GetInt(IDictionary<string, string?> cells, string column)
        {
            var text = GetString(cells, column);
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue))
            {
                return (int)decimalValue;
            }

            return 0;
        }

        private static decimal GetDecimal(IDictionary<string, string?> cells, string column)
        {
            var text = GetString(cells, column);
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            return 0;
        }

        private static decimal? GetNullableDecimal(IDictionary<string, string?> cells, string column)
        {
            var text = GetString(cells, column);
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }
             
            return null;
        }
    }
}
