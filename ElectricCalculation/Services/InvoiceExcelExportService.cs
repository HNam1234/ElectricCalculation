using System;
using System.Globalization;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

            // Template has a single sheet; use the first worksheet.
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

            var period = periodLabel?.Trim() ?? string.Empty;
            var issuer = issuerName?.Trim() ?? string.Empty;

            var householdName = customer.Name?.Trim() ?? string.Empty;
            var receiptNumber = customer.SequenceNumber > 0 ? customer.SequenceNumber : 1;

            var representativeName = customer.RepresentativeName?.Trim() ?? string.Empty;
            var groupName = customer.GroupName?.Trim() ?? string.Empty;
            var name = !string.IsNullOrWhiteSpace(representativeName) ? representativeName : householdName;
            var address = customer.Address?.Trim() ?? string.Empty;
            var location = customer.Location?.Trim() ?? string.Empty;
            var householdPhone = customer.HouseholdPhone?.Trim() ?? string.Empty;
            var representativePhone = customer.Phone?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(householdPhone) &&
                string.Equals(representativePhone, householdPhone, StringComparison.OrdinalIgnoreCase))
            {
                representativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(householdPhone) && !string.IsNullOrWhiteSpace(representativePhone))
            {
                householdPhone = representativePhone;
                representativePhone = string.Empty;
            }

            var meterNumber = customer.MeterNumber?.Trim() ?? string.Empty;
            var substation = customer.Substation?.Trim() ?? string.Empty;
            var bookCode = customer.BuildingName?.Trim() ?? string.Empty;
            var page = customer.Page?.Trim() ?? string.Empty;

            UpdateTextCell(sheetDataElement, mainNs, "I4", $"Số phiếu: {receiptNumber}");
            UpdateNumberCell(sheetDataElement, mainNs, "I6", receiptNumber);

            var periodText = FormatPeriodLabel(period);
            if (!string.IsNullOrWhiteSpace(periodText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
            }

            var recipient = !string.IsNullOrWhiteSpace(householdName) ? householdName : groupName;
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A5",
                string.IsNullOrWhiteSpace(recipient) ? string.Empty : $"Kính gửi: {recipient}");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A7",
                string.IsNullOrWhiteSpace(address) ? string.Empty : $"Địa chỉ hộ tiêu thụ: {address}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I7",
                string.IsNullOrWhiteSpace(householdPhone) ? string.Empty : $"Điện thoại: {householdPhone}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A8",
                string.IsNullOrWhiteSpace(name) ? string.Empty : $"Đại diện: {name}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I8",
                string.IsNullOrWhiteSpace(representativePhone) ? string.Empty : $"Điện thoại: {representativePhone}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I10",
                string.IsNullOrWhiteSpace(meterNumber) ? string.Empty : $"Số công tơ: {meterNumber}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I15",
                string.IsNullOrWhiteSpace(substation) ? string.Empty : $"TBA: {substation}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I16",
                string.IsNullOrWhiteSpace(bookCode) ? string.Empty : $"Mã sổ: {bookCode}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I17",
                string.IsNullOrWhiteSpace(page) ? string.Empty : $"Trang: {page}.");

            var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
            var consumption = customer.Consumption;
            var amount = customer.Amount;

            // B13/C13/D13/F13/G13: indexes, multiplier, subsidy, unit price.
            UpdateNumberCell(sheetDataElement, mainNs, "B13", customer.CurrentIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "C13", customer.PreviousIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "D13", multiplier);
            UpdateNumberCell(sheetDataElement, mainNs, "F13", customer.SubsidizedKwh);
            UpdateNumberCell(sheetDataElement, mainNs, "G13", customer.UnitPrice);

            // E13: consumption (kWh).
            UpdateNumberCell(sheetDataElement, mainNs, "E13", consumption);

            // H13: amount.
            UpdateNumberCell(sheetDataElement, mainNs, "H13", amount);
            UpdateNumberCell(sheetDataElement, mainNs, "H18", amount);

            // I13: meter location (optional).
            if (!string.IsNullOrWhiteSpace(location))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", $"Vị trí đặt: {location}.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", string.Empty);
            }

            // A19: amount in words.
            var amountText = VietnameseNumberTextService.ConvertAmountToText(amount);
            if (!string.IsNullOrWhiteSpace(amountText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A19", $"Bằng chữ: {amountText}./.");
            }

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "H21",
                $"Hà Nội, ngày {DateTime.Now.Day} tháng {DateTime.Now.Month} năm {DateTime.Now.Year}");

            UpdateTextCell(sheetDataElement, mainNs, "H27", issuer);

            using (var sheetWriteStream = sheetEntry.Open())
            {
                sheetWriteStream.SetLength(0);
                sheetDoc.Save(sheetWriteStream);
            }
        }

        public static void ExportInvoicesToWorkbook(
            string templatePath,
            string outputPath,
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
                throw new FileNotFoundException("Invoice template Excel file not found.", templatePath);
            }

            if (customers == null)
            {
                throw new ArgumentNullException(nameof(customers));
            }

            if (customers.Count == 0)
            {
                throw new ArgumentException("Customers list is empty.", nameof(customers));
            }

            // Copy template to output so we never modify the original file.
            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("Invoice template is invalid: missing xl/workbook.xml.");

            var relEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                ?? throw new InvalidOperationException("Invoice template is invalid: missing xl/_rels/workbook.xml.rels.");

            var contentTypesEntry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidOperationException("Invoice template is invalid: missing [Content_Types].xml.");

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
                ?? throw new InvalidOperationException("Invoice template is invalid: sheets collection not found.");

            var baseSheetElement = sheetsElement
                .Elements(mainNs + "sheet")
                .FirstOrDefault()
                ?? throw new InvalidOperationException("Invoice template is invalid: no worksheet found.");

            var relIdAttributeName = XName.Get("id", relNs.NamespaceName);
            var baseRelId = (string?)baseSheetElement.Attribute(relIdAttributeName);
            if (string.IsNullOrWhiteSpace(baseRelId))
            {
                throw new InvalidOperationException("Invoice template is invalid: worksheet has no relationship id.");
            }

            var relationshipsRoot = relDoc.Root
                ?? throw new InvalidOperationException("Invoice template is invalid: relationships root missing.");

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
            }

            var baseRelationship = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string?)r.Attribute("Id"), baseRelId, StringComparison.Ordinal));

            var target = (string?)baseRelationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                throw new InvalidOperationException("Invoice template is invalid: cannot locate worksheet content.");
            }

            var baseSheetPath = target.StartsWith("/", StringComparison.Ordinal)
                ? "xl" + target
                : "xl/" + target;

            var baseSheetEntry = archive.GetEntry(baseSheetPath)
                ?? throw new InvalidOperationException($"Invoice template is invalid: missing {baseSheetPath}.");

            var baseSheetBytes = ReadAllBytes(baseSheetEntry);

            var baseSheetIndex = TryParseWorksheetIndex(baseSheetPath);
            var baseSheetRelsBytes = baseSheetIndex > 0
                ? TryReadEntryBytes(archive, $"xl/worksheets/_rels/sheet{baseSheetIndex}.xml.rels")
                : null;

            var usedSheetIndexes = archive.Entries
                .Select(e => TryParseWorksheetIndex(e.FullName))
                .Where(i => i > 0)
                .ToHashSet();

            var maxSheetId = sheetsElement
                .Elements(mainNs + "sheet")
                .Select(s => (int?)s.Attribute("sheetId") ?? 0)
                .DefaultIfEmpty(0)
                .Max();

            var maxRelId = relationshipsRoot
                .Elements(relPackageNs + "Relationship")
                .Select(r => (string?)r.Attribute("Id"))
                .Where(id => id != null && Regex.IsMatch(id, @"^rId\d+$"))
                .Select(id => int.Parse(id![3..], CultureInfo.InvariantCulture))
                .DefaultIfEmpty(0)
                .Max();

            var usedSheetNames = sheetsElement
                .Elements(mainNs + "sheet")
                .Select(s => (string?)s.Attribute("name") ?? string.Empty)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < customers.Count; i++)
            {
                var customer = customers[i] ?? throw new ArgumentException("Customer is null.", nameof(customers));
                var sheetName = MakeUniqueSheetName(usedSheetNames, BuildSheetName(customer));

                ZipArchiveEntry sheetEntry;
                if (i == 0)
                {
                    baseSheetElement.SetAttributeValue("name", sheetName);
                    sheetEntry = baseSheetEntry;
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
                    WriteAllBytes(sheetEntry, baseSheetBytes);

                    if (baseSheetRelsBytes != null)
                    {
                        var relsEntry = archive.CreateEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels");
                        WriteAllBytes(relsEntry, baseSheetRelsBytes);
                    }

                    var relId = $"rId{++maxRelId}";
                    relationshipsRoot.Add(new XElement(relPackageNs + "Relationship",
                        new XAttribute("Id", relId),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", $"worksheets/sheet{sheetIndex}.xml")));

                    sheetsElement.Add(new XElement(mainNs + "sheet",
                        new XAttribute("name", sheetName),
                        new XAttribute("sheetId", ++maxSheetId),
                        new XAttribute(relIdAttributeName, relId)));

                    EnsureWorksheetContentType(contentTypesDoc, ctNs, sheetPath);
                }

                XDocument sheetDoc;
                using (var sheetReadStream = sheetEntry.Open())
                {
                    sheetDoc = XDocument.Load(sheetReadStream);
                }

                var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                    ?? throw new InvalidOperationException("Invoice template worksheet has no sheetData section.");

                PopulateInvoiceSheet(sheetDataElement, mainNs, customer, periodLabel, issuerName);

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
        }

        private static void PopulateInvoiceSheet(
            XElement sheetDataElement,
            XNamespace mainNs,
            Customer customer,
            string periodLabel,
            string issuerName)
        {
            var period = periodLabel?.Trim() ?? string.Empty;
            var issuer = issuerName?.Trim() ?? string.Empty;

            var householdName = customer.Name?.Trim() ?? string.Empty;
            var receiptNumber = customer.SequenceNumber > 0 ? customer.SequenceNumber : 1;

            var representativeName = customer.RepresentativeName?.Trim() ?? string.Empty;
            var groupName = customer.GroupName?.Trim() ?? string.Empty;
            var name = !string.IsNullOrWhiteSpace(representativeName) ? representativeName : householdName;
            var address = customer.Address?.Trim() ?? string.Empty;
            var location = customer.Location?.Trim() ?? string.Empty;
            var householdPhone = customer.HouseholdPhone?.Trim() ?? string.Empty;
            var representativePhone = customer.Phone?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(householdPhone) &&
                string.Equals(representativePhone, householdPhone, StringComparison.OrdinalIgnoreCase))
            {
                representativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(householdPhone) && !string.IsNullOrWhiteSpace(representativePhone))
            {
                householdPhone = representativePhone;
                representativePhone = string.Empty;
            }

            var meterNumber = customer.MeterNumber?.Trim() ?? string.Empty;
            var substation = customer.Substation?.Trim() ?? string.Empty;
            var bookCode = customer.BuildingName?.Trim() ?? string.Empty;
            var page = customer.Page?.Trim() ?? string.Empty;

            UpdateTextCell(sheetDataElement, mainNs, "I4", $"Số phiếu: {receiptNumber}");
            UpdateNumberCell(sheetDataElement, mainNs, "I6", receiptNumber);

            var periodText = FormatPeriodLabel(period);
            if (!string.IsNullOrWhiteSpace(periodText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
            }

            var recipient = !string.IsNullOrWhiteSpace(householdName) ? householdName : groupName;
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A5",
                string.IsNullOrWhiteSpace(recipient) ? string.Empty : $"Kính gửi: {recipient}");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A7",
                string.IsNullOrWhiteSpace(address) ? string.Empty : $"Địa chỉ hộ tiêu thụ: {address}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I7",
                string.IsNullOrWhiteSpace(householdPhone) ? string.Empty : $"Điện thoại: {householdPhone}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A8",
                string.IsNullOrWhiteSpace(name) ? string.Empty : $"Đại diện: {name}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I8",
                string.IsNullOrWhiteSpace(representativePhone) ? string.Empty : $"Điện thoại: {representativePhone}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I10",
                string.IsNullOrWhiteSpace(meterNumber) ? string.Empty : $"Số công tơ: {meterNumber}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I15",
                string.IsNullOrWhiteSpace(substation) ? string.Empty : $"TBA: {substation}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I16",
                string.IsNullOrWhiteSpace(bookCode) ? string.Empty : $"Mã sổ: {bookCode}.");

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I17",
                string.IsNullOrWhiteSpace(page) ? string.Empty : $"Trang: {page}.");

            var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
            var consumption = customer.Consumption;
            var amount = customer.Amount;

            UpdateNumberCell(sheetDataElement, mainNs, "B13", customer.CurrentIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "C13", customer.PreviousIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "D13", multiplier);
            UpdateNumberCell(sheetDataElement, mainNs, "F13", customer.SubsidizedKwh);
            UpdateNumberCell(sheetDataElement, mainNs, "G13", customer.UnitPrice);

            UpdateNumberCell(sheetDataElement, mainNs, "E13", consumption);
            UpdateNumberCell(sheetDataElement, mainNs, "H13", amount);
            UpdateNumberCell(sheetDataElement, mainNs, "H18", amount);

            if (!string.IsNullOrWhiteSpace(location))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", $"Vị trí đặt: {location}.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", string.Empty);
            }

            var amountText = VietnameseNumberTextService.ConvertAmountToText(amount);
            if (!string.IsNullOrWhiteSpace(amountText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "A19", $"Bằng chữ: {amountText}./.");
            }

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "H21",
                $"Hà Nội, ngày {DateTime.Now.Day} tháng {DateTime.Now.Month} năm {DateTime.Now.Year}");

            UpdateTextCell(sheetDataElement, mainNs, "H27", issuer);
        }

        private static void EnsureWorksheetContentType(XDocument contentTypesDoc, XNamespace ctNs, string sheetPath)
        {
            if (contentTypesDoc.Root == null)
            {
                return;
            }

            var partName = "/" + sheetPath.Replace("\\", "/");

            var exists = contentTypesDoc.Root
                .Elements(ctNs + "Override")
                .Any(e => string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));

            if (exists)
            {
                return;
            }

            contentTypesDoc.Root.Add(new XElement(ctNs + "Override",
                new XAttribute("PartName", partName),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
        }

        private static int TryParseWorksheetIndex(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return 0;
            }

            var match = Regex.Match(path.Replace("\\", "/"), @"/sheet(\d+)\.xml$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return 0;
            }

            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                ? index
                : 0;
        }

        private static byte[] ReadAllBytes(ZipArchiveEntry entry)
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        private static byte[]? TryReadEntryBytes(ZipArchive archive, string path)
        {
            var entry = archive.GetEntry(path);
            return entry == null ? null : ReadAllBytes(entry);
        }

        private static void WriteAllBytes(ZipArchiveEntry entry, byte[] bytes)
        {
            using var stream = entry.Open();
            stream.SetLength(0);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string BuildSheetName(Customer customer)
        {
            var seq = customer.SequenceNumber > 0 ? customer.SequenceNumber : 0;
            var name = string.IsNullOrWhiteSpace(customer.Name) ? "Khach" : customer.Name.Trim();
            var meter = string.IsNullOrWhiteSpace(customer.MeterNumber) ? string.Empty : $" - {customer.MeterNumber.Trim()}";
            var raw = $"{seq:0000} - {name}{meter}";

            raw = raw.Replace(":", " ")
                .Replace("\\", " ")
                .Replace("/", " ")
                .Replace("?", " ")
                .Replace("*", " ")
                .Replace("[", "(")
                .Replace("]", ")");

            raw = raw.Trim();
            if (raw.Length > 31)
            {
                raw = raw.Substring(0, 31).Trim();
            }

            return string.IsNullOrWhiteSpace(raw) ? $"HD {seq:0000}" : raw;
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
            cell.Elements(ns + "is").Remove();
            cell.Elements(ns + "v").Remove();

            if (value != null)
            {
                var vElement = new XElement(ns + "v", value.Value.ToString(CultureInfo.InvariantCulture));
                cell.Add(vElement);
            }

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

        private static string? FormatPeriodLabel(string periodLabel)
        {
            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return null;
            }

            var match = Regex.Match(periodLabel, @"(\d{1,2})\s*/\s*(\d{4})");
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var month) &&
                int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var year))
            {
                return month is >= 1 and <= 12 && year >= 2000
                    ? $"Tháng {month} năm {year}"
                    : periodLabel;
            }

            match = Regex.Match(periodLabel, @"tháng\s*(\d{1,2}).*?(\d{4})", RegexOptions.IgnoreCase);
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month) &&
                int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return month is >= 1 and <= 12 && year >= 2000
                    ? $"Tháng {month} năm {year}"
                    : periodLabel;
            }

            return periodLabel;
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

            string ReadThreeDigits(int number, bool isMostSignificantGroup)
            {
                int hundreds = number / 100;
                int tens = (number % 100) / 10;
                int ones = number % 10;

                var sb = new StringBuilder();

                if (hundreds > 0 || !isMostSignificantGroup)
                {
                    if (hundreds > 0)
                    {
                        sb.Append(unitNumbers[hundreds]).Append(" trăm");
                    }
                    else if (!isMostSignificantGroup)
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

            var groups = new List<int>(capacity: placeValues.Length);
            while (value > 0 && groups.Count < placeValues.Length)
            {
                groups.Add((int)(value % 1000));
                value /= 1000;
            }

            var highestGroupIndex = -1;
            for (var i = groups.Count - 1; i >= 0; i--)
            {
                if (groups[i] > 0)
                {
                    highestGroupIndex = i;
                    break;
                }
            }

            var resultBuilder = new StringBuilder();

            for (var groupIndex = highestGroupIndex; groupIndex >= 0; groupIndex--)
            {
                var groupNumber = groups[groupIndex];
                if (groupNumber <= 0)
                {
                    continue;
                }

                // Most significant group should not start with "không trăm" (e.g. 012 triệu => "mười hai triệu").
                var groupText = ReadThreeDigits(groupNumber, isMostSignificantGroup: groupIndex == highestGroupIndex);
                if (string.IsNullOrEmpty(groupText))
                {
                    continue;
                }

                if (resultBuilder.Length > 0)
                {
                    resultBuilder.Append(' ');
                }

                resultBuilder.Append(groupText);

                var unitText = placeValues[groupIndex];
                if (!string.IsNullOrEmpty(unitText))
                {
                    resultBuilder.Append(' ').Append(unitText);
                }
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
