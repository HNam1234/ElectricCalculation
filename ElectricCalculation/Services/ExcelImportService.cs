using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
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
        private enum ImportField
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

        private static string? AppendWarning(string? existing, string message)
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

            foreach (var pair in cells)
            {
                var normalized = NormalizeHeader(pair.Value);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                var field = GuessField(normalized);
                if (field == null)
                {
                    continue;
                }

                if (!map.ContainsKey(field.Value))
                {
                    map[field.Value] = pair.Key;
                }
            }

            var hasEnoughSignals =
                map.ContainsKey(ImportField.Name) &&
                (map.ContainsKey(ImportField.MeterNumber) ||
                 map.ContainsKey(ImportField.CurrentIndex) ||
                 map.ContainsKey(ImportField.PreviousIndex) ||
                 map.ContainsKey(ImportField.UnitPrice));

            if (!hasEnoughSignals)
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

        private static ImportField? GuessField(string normalizedHeader)
        {
            if (ContainsAny(normalizedHeader, "stt", "so thu tu", "so tt", "no."))
            {
                return ImportField.SequenceNumber;
            }

            if (ContainsAny(normalizedHeader, "ten khach", "khach hang", "ho ten", "ten"))
            {
                return ImportField.Name;
            }

            if (ContainsAny(normalizedHeader, "nhom", "don vi"))
            {
                return ImportField.GroupName;
            }

            if (ContainsAny(normalizedHeader, "loai"))
            {
                return ImportField.Category;
            }

            if (ContainsAny(normalizedHeader, "dia chi", "dc"))
            {
                return ImportField.Address;
            }

            if (ContainsAny(normalizedHeader, "dai dien", "nguoi dai dien"))
            {
                return ImportField.RepresentativeName;
            }

            if (ContainsAny(normalizedHeader, "dt ho", "dien thoai ho", "so dt ho"))
            {
                return ImportField.HouseholdPhone;
            }

            if (ContainsAny(normalizedHeader, "so cong to", "cong to", "meter"))
            {
                return ImportField.MeterNumber;
            }

            if (ContainsAny(normalizedHeader, "vi tri"))
            {
                return ImportField.Location;
            }

            if (ContainsAny(normalizedHeader, "tram", "tba"))
            {
                return ImportField.Substation;
            }

            if (ContainsAny(normalizedHeader, "trang", "page"))
            {
                return ImportField.Page;
            }

            if (ContainsAny(normalizedHeader, "nguoi th", "nguoi thu", "nguoi ghi"))
            {
                return ImportField.PerformedBy;
            }

            if (ContainsAny(normalizedHeader, "chi so", "cs"))
            {
                if (normalizedHeader.Contains("moi", StringComparison.OrdinalIgnoreCase))
                {
                    return ImportField.CurrentIndex;
                }

                if (normalizedHeader.Contains("cu", StringComparison.OrdinalIgnoreCase))
                {
                    return ImportField.PreviousIndex;
                }
            }

            if (ContainsAny(normalizedHeader, "he so", "hs", "multiplier"))
            {
                return ImportField.Multiplier;
            }

            if (ContainsAny(normalizedHeader, "don gia", "gia"))
            {
                return ImportField.UnitPrice;
            }

            if (ContainsAny(normalizedHeader, "bao cap", "tro cap", "mien giam"))
            {
                return ImportField.SubsidizedKwh;
            }

            if (ContainsAny(normalizedHeader, "so dt", "dien thoai", "dt"))
            {
                return ImportField.Phone;
            }

            if (ContainsAny(normalizedHeader, "nha", "toa", "ma so", "book"))
            {
                return ImportField.BuildingName;
            }

            return null;
        }

        private static bool ContainsAny(string haystack, params string[] needles)
        {
            if (string.IsNullOrWhiteSpace(haystack) || needles == null || needles.Length == 0)
            {
                return false;
            }

            foreach (var needle in needles)
            {
                if (string.IsNullOrWhiteSpace(needle))
                {
                    continue;
                }

                if (haystack.Contains(needle, StringComparison.OrdinalIgnoreCase))
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
