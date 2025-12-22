using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class InvoiceExcelExportService
    {
        public static void ExportInvoice(
            string templatePath,
            string outputPath,
            Customer customer,
            string periodLabel,
            string issuerName)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
            {
                throw new ArgumentException("Template path is required.", nameof(templatePath));
            }

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("Invoice template Excel file not found.", templatePath);
            }

            if (customer == null)
            {
                throw new ArgumentNullException(nameof(customer));
            }

            // Copy template to output so we never modify the original file.
            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("Invoice template is invalid: missing xl/workbook.xml.");

            XDocument workbookDoc;
            using (var workbookStream = workbookEntry.Open())
            {
                workbookDoc = XDocument.Load(workbookStream);
            }

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relPackageNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var sheetsElement = workbookDoc.Root?.Element(mainNs + "sheets")
                ?? throw new InvalidOperationException("Invoice template is invalid: sheets collection not found.");

            // Template chỉ có 1 sheet nên lấy sheet đầu tiên.
            var invoiceSheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault()
                ?? throw new InvalidOperationException("Invoice template is invalid: no worksheet found.");

            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);
            var relId = (string?)invoiceSheetElement.Attribute(relIdAttributeName);
            if (string.IsNullOrWhiteSpace(relId))
            {
                throw new InvalidOperationException("Invoice template is invalid: worksheet has no relationship id.");
            }

            var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                ?? throw new InvalidOperationException("Invoice template is invalid: missing xl/_rels/workbook.xml.rels.");

            XDocument relDoc;
            using (var relStream = relEntry.Open())
            {
                relDoc = XDocument.Load(relStream);
            }

            var relationshipsRoot = relDoc.Root
                ?? throw new InvalidOperationException("Invoice template is invalid: relationships root missing.");

            var relationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), relId, StringComparison.Ordinal));

            var target = (string?)relationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                throw new InvalidOperationException("Invoice template is invalid: cannot locate worksheet content.");
            }

            var sheetPath = target.StartsWith("/", StringComparison.Ordinal)
                ? "xl" + target
                : "xl/" + target;

            var sheetEntry = archive.GetEntry(sheetPath)
                ?? throw new InvalidOperationException($"Invoice template is invalid: missing {sheetPath}.");

            XDocument sheetDoc;
            using (var sheetReadStream = sheetEntry.Open())
            {
                sheetDoc = XDocument.Load(sheetReadStream);
            }

            var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                ?? throw new InvalidOperationException("Invoice template worksheet has no sheetData section.");

            // Xoá calcChain để Excel không cảnh báo "We found a problem with some content..."
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

            var period = periodLabel?.Trim() ?? string.Empty;
            var issuer = issuerName?.Trim() ?? string.Empty;

            var name = customer.Name?.Trim() ?? string.Empty;
            var groupName = customer.GroupName?.Trim() ?? string.Empty;
            var address = customer.Address?.Trim() ?? string.Empty;
            var location = customer.Location?.Trim() ?? string.Empty;
            var phone = customer.Phone?.Trim() ?? string.Empty;
            var meterNumber = customer.MeterNumber?.Trim() ?? string.Empty;

            // Header: kỳ, số phiếu, kênh gửi, địa chỉ, điện thoại, đại diện...
            UpdateTextCell(sheetDataElement, mainNs, "I4", "Số phiếu: 1");

            if (!string.IsNullOrWhiteSpace(period))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A2", $"Kỳ: {period}");
            }

            var channelTextSource = !string.IsNullOrWhiteSpace(groupName) ? groupName : name;
            if (!string.IsNullOrWhiteSpace(channelTextSource))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A5", $"Kênh gửi: {channelTextSource}");
            }

            if (!string.IsNullOrWhiteSpace(address))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A7", $"Địa chỉ hộ tiêu thụ: {address}.");
            }

            if (!string.IsNullOrWhiteSpace(phone))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I7", $"Điện thoại: {phone}.");
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A8", $"Đại diện: {name}.");
            }

            // Dòng I8 trong template là điện thoại người đại diện -> để trống cho đơn giản.
            UpdateTextCell(sheetDataElement, mainNs, "I8", string.Empty);

            if (!string.IsNullOrWhiteSpace(meterNumber))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I10", $"Số công tơ: {meterNumber}.");
            }

            var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
            var consumption = customer.Consumption;
            var amount = customer.Amount;

            // B13/C13/D13/F13/G13: chỉ số, hệ số, bao cấp, đơn giá
            UpdateNumberCell(sheetDataElement, mainNs, "B13", customer.CurrentIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "C13", customer.PreviousIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "D13", multiplier);
            UpdateNumberCell(sheetDataElement, mainNs, "F13", customer.SubsidizedKwh);
            UpdateNumberCell(sheetDataElement, mainNs, "G13", customer.UnitPrice);

            // E13: sản lượng (kWh)
            UpdateNumberCell(sheetDataElement, mainNs, "E13", consumption);

            // H13: thành tiền
            UpdateNumberCell(sheetDataElement, mainNs, "H13", amount);

            // I13: Vị trí đặt công tơ (nếu có)
            if (!string.IsNullOrWhiteSpace(location))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", $"Vị trí đặt: {location}.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", string.Empty);
            }

            // A19: số tiền bằng chữ
            var amountText = ConvertAmountToVietnameseText(amount);
            if (!string.IsNullOrWhiteSpace(amountText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A19", $"Bằng chữ: {amountText}./.");
            }

            // H21: Người lập đơn
            if (!string.IsNullOrWhiteSpace(issuer))
            {
                UpdateTextCell(sheetDataElement, mainNs, "H21", issuer);
            }

            using (var sheetWriteStream = sheetEntry.Open())
            {
                sheetWriteStream.SetLength(0);
                sheetDoc.Save(sheetWriteStream);
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
            var isElement = new XElement(ns + "is", new XElement(ns + "t", text));
            cell.Add(isElement);

            if (!string.IsNullOrEmpty(styleAttr))
            {
                cell.SetAttributeValue("s", styleAttr);
            }
        }

        private static void UpdateNumberCell(XElement sheetDataElement, XNamespace ns, string cellReference, decimal value)
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
            cell.Elements(ns + "is").Remove();
            cell.Elements(ns + "v").Remove();

            var vElement = new XElement(ns + "v", value.ToString(CultureInfo.InvariantCulture));
            cell.Add(vElement);

            if (!string.IsNullOrEmpty(styleAttr))
            {
                cell.SetAttributeValue("s", styleAttr);
            }
        }

        private static int GetRowIndex(string cellReference)
        {
            var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex)
                ? rowIndex
                : 0;
        }

        private static string ConvertAmountToVietnameseText(decimal amount)
        {
            var rounded = Math.Round(amount, 0, MidpointRounding.AwayFromZero);
            if (rounded <= 0)
            {
                return "Không đồng";
            }

            var value = (long)rounded;
            if (value < 0)
            {
                value = -value;
            }

            string[] unitNumbers =
            {
                "không", "một", "hai", "ba", "bốn",
                "năm", "sáu", "bảy", "tám", "chín"
            };

            string[] placeValues =
            {
                string.Empty,
                "nghìn",
                "triệu",
                "tỷ",
                "nghìn tỷ",
                "triệu tỷ"
            };

            string ReadThreeDigits(int number, bool isFirstGroup)
            {
                int hundreds = number / 100;
                int tens = (number % 100) / 10;
                int ones = number % 10;

                var sb = new StringBuilder();

                if (hundreds > 0 || !isFirstGroup)
                {
                    if (hundreds > 0)
                    {
                        sb.Append(unitNumbers[hundreds]).Append(" trăm");
                    }
                    else if (!isFirstGroup)
                    {
                        sb.Append("không trăm");
                    }
                }

                if (tens > 1)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(' ');
                    }

                    sb.Append(unitNumbers[tens]).Append(" mươi");

                    if (ones == 1)
                    {
                        sb.Append(" mốt");
                    }
                    else if (ones == 4)
                    {
                        sb.Append(" tư");
                    }
                    else if (ones == 5)
                    {
                        sb.Append(" lăm");
                    }
                    else if (ones > 0)
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }
                else if (tens == 1)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(' ');
                    }

                    sb.Append("mười");

                    if (ones == 1)
                    {
                        sb.Append(" một");
                    }
                    else if (ones == 4)
                    {
                        sb.Append(" bốn");
                    }
                    else if (ones == 5)
                    {
                        sb.Append(" lăm");
                    }
                    else if (ones > 0)
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }
                else if (tens == 0 && ones > 0)
                {
                    if (sb.Length > 0)
                    {
                        sb.Append(" lẻ");
                    }

                    if (ones == 5 && sb.Length > 0)
                    {
                        sb.Append(" năm");
                    }
                    else
                    {
                        sb.Append(' ').Append(unitNumbers[ones]);
                    }
                }

                return sb.ToString().Trim();
            }

            var resultBuilder = new StringBuilder();
            var groupIndex = 0;

            while (value > 0 && groupIndex < placeValues.Length)
            {
                var groupNumber = (int)(value % 1000);
                if (groupNumber > 0)
                {
                    var groupText = ReadThreeDigits(groupNumber, value < 1000 && groupIndex == 0);
                    if (!string.IsNullOrEmpty(groupText))
                    {
                        var unitText = placeValues[groupIndex];
                        if (resultBuilder.Length == 0)
                        {
                            resultBuilder.Insert(0, groupText + (string.IsNullOrEmpty(unitText) ? string.Empty : " " + unitText));
                        }
                        else
                        {
                            resultBuilder.Insert(0, groupText + (string.IsNullOrEmpty(unitText) ? " " : " " + unitText + " "));
                        }
                    }
                }

                value /= 1000;
                groupIndex++;
            }

            var result = resultBuilder.ToString().Trim();
            if (result.Length == 0)
            {
                result = "không";
            }

            result = char.ToUpper(result[0], CultureInfo.CurrentCulture) + result[1..] + " đồng";

            return result;
        }
    }
}

