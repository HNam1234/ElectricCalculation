using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    // Simple importer for the summary template:
    // Prefer sheet "Data" (fallback to a matching sheet), data starts at row 5, key columns:
    // A: No., B: Customer, C: Group, D: Address, E: Household phone, F: Representative,
    // G: Rep phone, H: Building, J: Meter, K: Category, L: Location, M: Substation, N: Page,
    // O: Current, P: Previous, Q: Multiplier, S: Subsidy, U: Unit price, W: Performed by.
    public static class ExcelImportService
    {
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
                "Bang ke",
                "Ban  in so",
                "Ban in so"
            };

            var dataSheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault(s =>
                {
                    var name = (string?)s.Attribute("name");
                    return !string.IsNullOrWhiteSpace(name) &&
                           preferredSheetNames.Any(preferred => string.Equals(name, preferred, StringComparison.OrdinalIgnoreCase));
                }) ?? sheetsElement.Elements(mainNs + "sheet").FirstOrDefault();

            var selectedSheetName = (string?)dataSheetElement?.Attribute("name") ?? "Data";
            if (!string.Equals(selectedSheetName, "Data", StringComparison.OrdinalIgnoreCase))
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

            foreach (var row in sheetDataElement.Elements(mainNs + "row"))
            {
                if (!int.TryParse((string?)row.Attribute("r"), out var rowIndex))
                {
                    continue;
                }

                // Skip header; data starts at row 5.
                if (rowIndex < 5)
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

                // Skip rows without sequence number.
                var sequenceNumber = GetInt(cells, "A");
                if (sequenceNumber <= 0)
                {
                    continue;
                }

                var householdPhone = GetString(cells, "E") ?? string.Empty;
                var representativePhone = GetString(cells, "G") ?? string.Empty;

                var customer = new Customer
                {
                    SequenceNumber = sequenceNumber,
                    Name = GetString(cells, "B") ?? string.Empty,
                    GroupName = GetString(cells, "C") ?? string.Empty,
                    Address = GetString(cells, "D") ?? string.Empty,
                    RepresentativeName = GetString(cells, "F") ?? string.Empty,
                    HouseholdPhone = householdPhone,
                    Phone = !string.IsNullOrWhiteSpace(representativePhone) ? representativePhone : householdPhone,
                    BuildingName = GetString(cells, "H") ?? string.Empty,
                    MeterNumber = GetString(cells, "J") ?? string.Empty,
                    Category = GetString(cells, "K") ?? string.Empty,
                    Location = GetString(cells, "L") ?? string.Empty
                };

                customer.Substation = GetString(cells, "M") ?? string.Empty;
                customer.Page = GetString(cells, "N") ?? string.Empty;

                customer.CurrentIndex = GetDecimal(cells, "O");
                customer.PreviousIndex = GetDecimal(cells, "P");
                customer.Multiplier = GetDecimal(cells, "Q");
                customer.SubsidizedKwh = GetDecimal(cells, "S");
                customer.UnitPrice = GetDecimal(cells, "U");

                customer.PerformedBy = GetString(cells, "W") ?? string.Empty;

                if (customer.Multiplier <= 0)
                {
                    customer.Multiplier = 1;
                }

                result.Add(customer);
            }

            return result;
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
    }
}
