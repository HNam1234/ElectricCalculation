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
using ElectricCalculation.Helpers;
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
            var visibleZeroStyleResolver = VisibleZeroStyleResolver.TryLoad(archive);

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

            if (sheetDoc.Root != null)
            {
                EnsureWorksheetShowsZeros(sheetDoc.Root, mainNs);
                EnsureInvoiceColumnWidths(sheetDoc.Root, mainNs);
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
            var groupName = TextNormalization.NormalizeForDisplay(customer.GroupName);
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
            var hasValidReading = !customer.IsMissingReading && !customer.HasReadingError;

            // B13/C13/D13/F13/G13: indexes, multiplier, subsidy, unit price.
            UpdateNumberCell(sheetDataElement, mainNs, "B13", customer.CurrentIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "C13", customer.PreviousIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "D13", multiplier);
            UpdateNumberCell(sheetDataElement, mainNs, "F13", customer.SubsidizedKwh);
            UpdateNumberCell(sheetDataElement, mainNs, "G13", customer.UnitPrice);

            var worksheetRoot = sheetDoc.Root;

            // E13: consumption (kWh). Show 0 explicitly for zero-usage rows (template may hide numeric zeros).
            if (hasValidReading)
            {
                UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, "E13", consumption, visibleZeroStyleResolver);
            }
            else
            {
                UpdateNumberCell(sheetDataElement, mainNs, "E13", null);
            }

            // H13/H18: amount. Keep blank for missing/error readings; show 0 explicitly when applicable.
            if (hasValidReading)
            {
                UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, "H13", amount, visibleZeroStyleResolver);
                UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, "H18", amount, visibleZeroStyleResolver);
            }
            else
            {
                UpdateNumberCell(sheetDataElement, mainNs, "H13", null);
                UpdateNumberCell(sheetDataElement, mainNs, "H18", null);
            }

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
            if (hasValidReading)
            {
                var amountText = VietnameseNumberTextService.ConvertAmountToText(amount);
                UpdateTextCell(
                    sheetDataElement,
                    mainNs,
                    "A19",
                    string.IsNullOrWhiteSpace(amountText) ? string.Empty : $"Bằng chữ: {amountText}./.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "A19", string.Empty);
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

            visibleZeroStyleResolver?.SaveIfDirty();
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
            var visibleZeroStyleResolver = VisibleZeroStyleResolver.TryLoad(archive);

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

                if (sheetDoc.Root != null)
                {
                    EnsureWorksheetShowsZeros(sheetDoc.Root, mainNs);
                    EnsureInvoiceColumnWidths(sheetDoc.Root, mainNs);
                }

                var sheetDataElement = sheetDoc.Root?.Element(mainNs + "sheetData")
                    ?? throw new InvalidOperationException("Invoice template worksheet has no sheetData section.");

                PopulateInvoiceSheet(sheetDataElement, mainNs, customer, periodLabel, issuerName, visibleZeroStyleResolver);

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

            visibleZeroStyleResolver?.SaveIfDirty();
        }

        public static void ExportHouseholdsAsSingleInvoice(
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

            var householdList = customers
                .Where(c => c != null)
                .ToList();

            if (householdList.Count == 0)
            {
                throw new ArgumentException("Customers list is empty.", nameof(customers));
            }

            // Copy template to output so we never modify the original file.
            File.Copy(templatePath, outputPath, overwrite: true);

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            var visibleZeroStyleResolver = VisibleZeroStyleResolver.TryLoad(archive);

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

            XDocument sheetDoc;
            using (var sheetReadStream = sheetEntry.Open())
            {
                sheetDoc = XDocument.Load(sheetReadStream);
            }

            var worksheetRoot = sheetDoc.Root
                ?? throw new InvalidOperationException("Invoice template worksheet is empty.");

            EnsureWorksheetShowsZeros(worksheetRoot, mainNs);
            EnsureInvoiceColumnWidths(worksheetRoot, mainNs);

            var sheetDataElement = worksheetRoot.Element(mainNs + "sheetData")
                ?? throw new InvalidOperationException("Invoice template worksheet has no sheetData section.");

            PopulateMultiHouseholdInvoiceSheet(
                sheetDataElement,
                worksheetRoot,
                mainNs,
                householdList,
                periodLabel,
                issuerName,
                visibleZeroStyleResolver);

            using (var sheetWriteStream = sheetEntry.Open())
            {
                sheetWriteStream.SetLength(0);
                sheetDoc.Save(sheetWriteStream);
            }

            visibleZeroStyleResolver?.SaveIfDirty();
        }

        private static void PopulateMultiHouseholdInvoiceSheet(
            XElement sheetDataElement,
            XElement worksheetRoot,
            XNamespace mainNs,
            IReadOnlyList<Customer> customers,
            string periodLabel,
            string issuerName,
            VisibleZeroStyleResolver? visibleZeroStyleResolver)
        {
            if (customers.Count == 1)
            {
                PopulateInvoiceSheet(sheetDataElement, mainNs, customers[0], periodLabel, issuerName, visibleZeroStyleResolver);
                return;
            }

            const int detailStartRow = 13;
            const int detailHeaderRow = 10;
            const int rowsAfterDetailStart = 14;
            const int baseTotalRow = 18;
            const int baseAmountInWordsRow = 19;
            const int baseDateRow = 21;
            const int baseIssuerRow = 27;
            const int baseSubstationRow = 15;
            const int baseBookCodeRow = 16;
            const int basePageRow = 17;

            var extraRows = customers.Count - 1;
            var detailTemplateRow = GetRow(sheetDataElement, mainNs, detailStartRow)
                ?? throw new InvalidOperationException("Invoice template is missing detail row 13.");

            ConfigureMultiHouseholdColumns(worksheetRoot, mainNs);
            EnsureInvoiceColumnWidths(worksheetRoot, mainNs);
            RemoveAmountColumnMergedRange(worksheetRoot, mainNs);
            RemoveMergeRange(worksheetRoot, mainNs, "J10:K12");
            RemoveMergeRange(worksheetRoot, mainNs, "J13:K14");
            RemoveMergeRange(worksheetRoot, mainNs, "K10:L12");
            RemoveMergeRange(worksheetRoot, mainNs, "K13:L14");
            RemoveMergeRange(worksheetRoot, mainNs, $"J{detailHeaderRow}:J{detailHeaderRow + 2}");

            if (extraRows > 0)
            {
                ShiftRows(sheetDataElement, mainNs, rowsAfterDetailStart, extraRows);
                ShiftMergedRanges(worksheetRoot, mainNs, rowsAfterDetailStart, extraRows);
            }

            EnsureDetailRows(sheetDataElement, mainNs, detailTemplateRow, customers.Count, detailStartRow);
            var displayNames = BuildMultiHouseholdDisplayNames(customers);

            // Multi-household layout:
            // A: STT, B: Tên khách, C..I: chỉ số và tiền.
            UpdateTextCell(sheetDataElement, mainNs, $"B{detailHeaderRow}", "Tên khách");
            UpdateTextCell(sheetDataElement, mainNs, $"C{detailHeaderRow}", "Chỉ số mới");
            UpdateTextCell(sheetDataElement, mainNs, $"D{detailHeaderRow}", "Chỉ số cũ");
            UpdateTextCell(sheetDataElement, mainNs, $"E{detailHeaderRow}", "Hệ số");
            UpdateTextCell(sheetDataElement, mainNs, $"F{detailHeaderRow}", "Điện năng tiêu thụ (kWh)");
            UpdateTextCell(sheetDataElement, mainNs, $"G{detailHeaderRow}", "Bao cấp (kWh)");
            UpdateTextCell(sheetDataElement, mainNs, $"H{detailHeaderRow}", "Đơn giá (VNĐ)");
            UpdateTextCell(sheetDataElement, mainNs, $"I{detailHeaderRow}", "Thành tiền (VNĐ)");
            CopyCellStyle(sheetDataElement, mainNs, $"H{detailHeaderRow}", $"I{detailHeaderRow}");

            for (var i = 0; i < customers.Count; i++)
            {
                var rowIndex = detailStartRow + i;
                var customer = customers[i];
                var sequence = i + 1;
                var multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
                var consumption = customer.Consumption;
                var amount = customer.Amount;
                var hasValidReading = !customer.IsMissingReading && !customer.HasReadingError;
                var displayName = displayNames[i];

                UpdateNumberCell(sheetDataElement, mainNs, $"A{rowIndex}", sequence);
                UpdateTextCell(sheetDataElement, mainNs, $"B{rowIndex}", displayName);
                ApplyShrinkToFitTextStyle(sheetDataElement, worksheetRoot, mainNs, $"B{rowIndex}", visibleZeroStyleResolver);
                UpdateNumberCell(sheetDataElement, mainNs, $"C{rowIndex}", customer.CurrentIndex);
                UpdateNumberCell(sheetDataElement, mainNs, $"D{rowIndex}", customer.PreviousIndex);
                UpdateNumberCell(sheetDataElement, mainNs, $"E{rowIndex}", multiplier);
                if (hasValidReading)
                {
                    UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, $"F{rowIndex}", consumption, visibleZeroStyleResolver);
                }
                else
                {
                    UpdateNumberCell(sheetDataElement, mainNs, $"F{rowIndex}", null);
                }
                UpdateNumberCell(sheetDataElement, mainNs, $"G{rowIndex}", customer.SubsidizedKwh);
                UpdateNumberCell(sheetDataElement, mainNs, $"H{rowIndex}", customer.UnitPrice);
                if (hasValidReading)
                {
                    UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, $"I{rowIndex}", amount, visibleZeroStyleResolver);
                }
                else
                {
                    UpdateNumberCell(sheetDataElement, mainNs, $"I{rowIndex}", null);
                }
                CopyCellStyle(sheetDataElement, mainNs, $"H{rowIndex}", $"I{rowIndex}");
            }

            var totalRow = baseTotalRow + extraRows;
            var amountInWordsRow = baseAmountInWordsRow + extraRows;
            var dateRow = baseDateRow + extraRows;
            var issuerRow = baseIssuerRow + extraRows;
            var substationRow = baseSubstationRow + extraRows;
            var bookCodeRow = baseBookCodeRow + extraRows;
            var pageRow = basePageRow + extraRows;

            ReplaceMergeRange(
                worksheetRoot,
                mainNs,
                $"A{amountInWordsRow}:H{amountInWordsRow}",
                $"A{amountInWordsRow}:I{amountInWordsRow}");

            var groupName = TextNormalization.NormalizeForDisplay(GetSharedNonEmptyValue(customers, c => c.GroupName));
            if (string.IsNullOrWhiteSpace(groupName))
            {
                groupName = "Nhiều hộ";
            }
            var sharedAddress = GetSharedNonEmptyValue(customers, c => c.Address) ?? string.Empty;
            var sharedRepresentative = GetSharedNonEmptyValue(customers, c => c.RepresentativeName) ?? string.Empty;
            var sharedHouseholdPhone = GetSharedNonEmptyValue(customers, c => c.HouseholdPhone) ?? string.Empty;
            var sharedRepresentativePhone = GetSharedNonEmptyValue(customers, c => c.Phone) ?? string.Empty;
            var sharedSubstation = GetSharedNonEmptyValue(customers, c => c.Substation) ?? string.Empty;
            var sharedBookCode = GetSharedNonEmptyValue(customers, c => c.BuildingName) ?? string.Empty;
            var sharedPage = GetSharedNonEmptyValue(customers, c => c.Page) ?? string.Empty;
            var issuer = issuerName?.Trim() ?? string.Empty;
            var representativeDisplay = !string.IsNullOrWhiteSpace(sharedRepresentative) ? sharedRepresentative : groupName;

            UpdateTextCell(sheetDataElement, mainNs, "I4", string.Empty);
            UpdateNumberCell(sheetDataElement, mainNs, "I6", null);
            UpdateTextCell(sheetDataElement, mainNs, "J4", string.Empty);
            UpdateNumberCell(sheetDataElement, mainNs, "J6", null);
            UpdateTextCell(sheetDataElement, mainNs, "J7", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, "J8", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"J{substationRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"J{bookCodeRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"J{pageRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, "K4", string.Empty);
            UpdateNumberCell(sheetDataElement, mainNs, "K6", null);
            UpdateTextCell(sheetDataElement, mainNs, "K7", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, "K8", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, "K10", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"K{substationRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"K{bookCodeRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"K{pageRow}", string.Empty);

            if (!string.IsNullOrWhiteSpace(sharedHouseholdPhone) &&
                string.Equals(sharedRepresentativePhone, sharedHouseholdPhone, StringComparison.OrdinalIgnoreCase))
            {
                sharedRepresentativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(sharedHouseholdPhone) && !string.IsNullOrWhiteSpace(sharedRepresentativePhone))
            {
                sharedHouseholdPhone = sharedRepresentativePhone;
                sharedRepresentativePhone = string.Empty;
            }

            var periodText = FormatPeriodLabel(periodLabel);
            if (!string.IsNullOrWhiteSpace(periodText))
            {
                UpdateTextCell(sheetDataElement, mainNs, "F2", periodText);
            }

            UpdateTextCell(sheetDataElement, mainNs, "A5", $"Kính gửi: {groupName}");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A7",
                string.IsNullOrWhiteSpace(sharedAddress)
                    ? string.Empty
                    : $"Địa chỉ hộ tiêu thụ: {sharedAddress}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I7",
                string.IsNullOrWhiteSpace(sharedHouseholdPhone) ? string.Empty : $"Điện thoại: {sharedHouseholdPhone}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "A8",
                string.IsNullOrWhiteSpace(representativeDisplay) ? string.Empty : $"Đại diện: {representativeDisplay}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                "I8",
                string.IsNullOrWhiteSpace(sharedRepresentativePhone) ? string.Empty : $"Điện thoại: {sharedRepresentativePhone}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"I{substationRow}",
                string.IsNullOrWhiteSpace(sharedSubstation) ? string.Empty : $"TBA: {sharedSubstation}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"I{bookCodeRow}",
                string.IsNullOrWhiteSpace(sharedBookCode) ? string.Empty : $"Mã sổ: {sharedBookCode}.");
            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"I{pageRow}",
                string.IsNullOrWhiteSpace(sharedPage) ? string.Empty : $"Trang: {sharedPage}.");

            var totalAmount = customers.Sum(c => c.Amount);
            CopyCellStyle(sheetDataElement, mainNs, $"H{totalRow}", $"I{totalRow}");
            CopyCellStyle(sheetDataElement, mainNs, $"G{totalRow}", $"H{totalRow}");
            UpdateTextCell(sheetDataElement, mainNs, $"G{totalRow}", string.Empty);
            UpdateTextCell(sheetDataElement, mainNs, $"H{totalRow}", "Tổng cộng:");
            UpdateNumberCellVisibleZero(sheetDataElement, worksheetRoot, mainNs, $"I{totalRow}", totalAmount, visibleZeroStyleResolver);

            var amountText = VietnameseNumberTextService.ConvertAmountToText(totalAmount);
            if (!string.IsNullOrWhiteSpace(amountText))
            {
                UpdateTextCell(sheetDataElement, mainNs, $"A{amountInWordsRow}", $"Bằng chữ: {amountText}./.");
            }

            UpdateTextCell(
                sheetDataElement,
                mainNs,
                $"H{dateRow}",
                $"Hà Nội, ngày {DateTime.Now.Day} tháng {DateTime.Now.Month} năm {DateTime.Now.Year}");
            ReplaceMergeRange(worksheetRoot, mainNs, $"H{dateRow}:I{dateRow}", $"H{dateRow}:J{dateRow}");
            ReplaceMergeRange(worksheetRoot, mainNs, $"H{dateRow + 1}:I{dateRow + 1}", $"H{dateRow + 1}:J{dateRow + 1}");
            ReplaceMergeRange(worksheetRoot, mainNs, $"H{issuerRow}:I{issuerRow}", $"H{issuerRow}:J{issuerRow}");

            UpdateTextCell(sheetDataElement, mainNs, $"H{issuerRow}", issuer);
        }

        private static string? GetSharedNonEmptyValue(
            IReadOnlyList<Customer> customers,
            Func<Customer, string?> selector)
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
                .Select((customer, index) => GetMultiHouseholdBaseName(customer, index + 1))
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

        private static string GetMultiHouseholdBaseName(Customer customer, int sequence)
        {
            return string.IsNullOrWhiteSpace(customer.Name)
                ? $"Hộ {sequence}"
                : customer.Name.Trim();
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

        private static XElement? GetRow(XElement sheetDataElement, XNamespace mainNs, int rowIndex)
        {
            return sheetDataElement
                .Elements(mainNs + "row")
                .FirstOrDefault(r =>
                    string.Equals(
                        (string?)r.Attribute("r"),
                        rowIndex.ToString(CultureInfo.InvariantCulture),
                        StringComparison.Ordinal));
        }

        private static void EnsureDetailRows(
            XElement sheetDataElement,
            XNamespace mainNs,
            XElement templateRow,
            int rowCount,
            int startRowIndex)
        {
            for (var i = 1; i < rowCount; i++)
            {
                var targetRowIndex = startRowIndex + i;
                if (GetRow(sheetDataElement, mainNs, targetRowIndex) != null)
                {
                    continue;
                }

                var clone = new XElement(templateRow);
                ReindexRowAndCells(clone, mainNs, targetRowIndex);
                InsertRowInOrder(sheetDataElement, mainNs, clone);
            }
        }

        private static void ShiftRows(XElement sheetDataElement, XNamespace mainNs, int fromRowIndex, int delta)
        {
            if (delta <= 0)
            {
                return;
            }

            var rowsToShift = sheetDataElement
                .Elements(mainNs + "row")
                .Select(r => new
                {
                    Row = r,
                    Index = int.TryParse((string?)r.Attribute("r"), out var rowIndex) ? rowIndex : 0
                })
                .Where(x => x.Index >= fromRowIndex)
                .OrderByDescending(x => x.Index)
                .ToList();

            foreach (var entry in rowsToShift)
            {
                var row = entry.Row;
                var originalRowIndex = entry.Index;
                var shiftedRowIndex = originalRowIndex + delta;
                row.SetAttributeValue("r", shiftedRowIndex.ToString(CultureInfo.InvariantCulture));

                foreach (var cell in row.Elements(mainNs + "c"))
                {
                    var cellRef = (string?)cell.Attribute("r");
                    if (string.IsNullOrWhiteSpace(cellRef))
                    {
                        continue;
                    }

                    cell.SetAttributeValue("r", ShiftCellReference(cellRef, delta));
                }
            }
        }

        private static void ShiftMergedRanges(XElement worksheetRoot, XNamespace mainNs, int fromRowIndex, int delta)
        {
            if (delta <= 0)
            {
                return;
            }

            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells == null)
            {
                return;
            }

            foreach (var mergeCell in mergeCells.Elements(mainNs + "mergeCell"))
            {
                var mergeRef = (string?)mergeCell.Attribute("ref");
                if (string.IsNullOrWhiteSpace(mergeRef))
                {
                    continue;
                }

                if (!TryParseRange(mergeRef, out var startColumn, out var startRow, out var endColumn, out var endRow))
                {
                    continue;
                }

                if (endRow < fromRowIndex)
                {
                    continue;
                }

                if (startRow >= fromRowIndex)
                {
                    startRow += delta;
                    endRow += delta;
                }
                else
                {
                    endRow += delta;
                }

                mergeCell.SetAttributeValue("ref", $"{startColumn}{startRow}:{endColumn}{endRow}");
            }
        }

        private static void ConfigureMultiHouseholdColumns(XElement worksheetRoot, XNamespace mainNs)
        {
            var cols = worksheetRoot.Element(mainNs + "cols");
            if (cols == null)
            {
                cols = new XElement(mainNs + "cols");
                var sheetData = worksheetRoot.Element(mainNs + "sheetData");
                if (sheetData != null)
                {
                    sheetData.AddBeforeSelf(cols);
                }
                else
                {
                    worksheetRoot.Add(cols);
                }
            }

            cols.RemoveNodes();

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 1),
                new XAttribute("max", 1),
                new XAttribute("width", "4.7109375"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 2),
                new XAttribute("max", 2),
                new XAttribute("width", "48"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 3),
                new XAttribute("max", 3),
                new XAttribute("width", "10.42578125"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 4),
                new XAttribute("max", 4),
                new XAttribute("width", "10.42578125"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 5),
                new XAttribute("max", 5),
                new XAttribute("width", "7.28515625"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 6),
                new XAttribute("max", 6),
                new XAttribute("width", "10"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 7),
                new XAttribute("max", 7),
                new XAttribute("width", "12"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 8),
                new XAttribute("max", 8),
                new XAttribute("width", "10.42578125"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 9),
                new XAttribute("max", 9),
                new XAttribute("width", "20.7109375"),
                new XAttribute("customWidth", 1)));

            cols.Add(new XElement(mainNs + "col",
                new XAttribute("min", 10),
                new XAttribute("max", 16384),
                new XAttribute("width", "7.140625"),
                new XAttribute("customWidth", 1)));
        }

        private static void RemoveAmountColumnMergedRange(XElement worksheetRoot, XNamespace mainNs)
        {
            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells == null)
            {
                return;
            }

            var toRemove = mergeCells
                .Elements(mainNs + "mergeCell")
                .Where(m =>
                {
                    var mergeRef = (string?)m.Attribute("ref");
                    if (string.IsNullOrWhiteSpace(mergeRef))
                    {
                        return false;
                    }

                    if (!TryParseRange(mergeRef, out var startColumn, out var startRow, out var endColumn, out var endRow))
                    {
                        return false;
                    }

                    return string.Equals(startColumn, "I", StringComparison.OrdinalIgnoreCase) &&
                           string.Equals(endColumn, "I", StringComparison.OrdinalIgnoreCase) &&
                           startRow == 13 &&
                           endRow >= 14;
                })
                .ToList();

            foreach (var mergeCell in toRemove)
            {
                mergeCell.Remove();
            }

            UpdateMergeCellsCount(mergeCells);
        }

        private static void EnsureMergeRange(XElement worksheetRoot, XNamespace mainNs, string mergeReference)
        {
            if (string.IsNullOrWhiteSpace(mergeReference))
            {
                return;
            }

            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells == null)
            {
                mergeCells = new XElement(mainNs + "mergeCells");
                var pageMargins = worksheetRoot.Element(mainNs + "pageMargins");
                if (pageMargins != null)
                {
                    pageMargins.AddBeforeSelf(mergeCells);
                }
                else
                {
                    worksheetRoot.Add(mergeCells);
                }
            }

            var exists = mergeCells
                .Elements(mainNs + "mergeCell")
                .Any(m => string.Equals((string?)m.Attribute("ref"), mergeReference, StringComparison.OrdinalIgnoreCase));

            if (!exists)
            {
                mergeCells.Add(new XElement(mainNs + "mergeCell", new XAttribute("ref", mergeReference)));
            }

            UpdateMergeCellsCount(mergeCells);
        }

        private static void RemoveMergeRange(XElement worksheetRoot, XNamespace mainNs, string mergeReference)
        {
            if (string.IsNullOrWhiteSpace(mergeReference))
            {
                return;
            }

            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells == null)
            {
                return;
            }

            var items = mergeCells
                .Elements(mainNs + "mergeCell")
                .Where(m => string.Equals((string?)m.Attribute("ref"), mergeReference, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var item in items)
            {
                item.Remove();
            }

            UpdateMergeCellsCount(mergeCells);
        }

        private static void ReplaceMergeRange(
            XElement worksheetRoot,
            XNamespace mainNs,
            string oldReference,
            string newReference)
        {
            if (string.IsNullOrWhiteSpace(newReference))
            {
                return;
            }

            var mergeCells = worksheetRoot.Element(mainNs + "mergeCells");
            if (mergeCells != null && !string.IsNullOrWhiteSpace(oldReference))
            {
                var oldItems = mergeCells
                    .Elements(mainNs + "mergeCell")
                    .Where(m => string.Equals((string?)m.Attribute("ref"), oldReference, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var item in oldItems)
                {
                    item.Remove();
                }

                UpdateMergeCellsCount(mergeCells);
            }

            EnsureMergeRange(worksheetRoot, mainNs, newReference);
        }

        private static void UpdateMergeCellsCount(XElement mergeCells)
        {
            var count = mergeCells.Elements().Count();
            if (count > 0)
            {
                mergeCells.SetAttributeValue("count", count);
            }
            else
            {
                mergeCells.Remove();
            }
        }

        private static void ReindexRowAndCells(XElement row, XNamespace mainNs, int newRowIndex)
        {
            row.SetAttributeValue("r", newRowIndex.ToString(CultureInfo.InvariantCulture));

            foreach (var cell in row.Elements(mainNs + "c"))
            {
                var cellRef = (string?)cell.Attribute("r");
                if (!string.IsNullOrWhiteSpace(cellRef))
                {
                    cell.SetAttributeValue("r", ReplaceCellRow(cellRef, newRowIndex));
                }

                cell.Elements(mainNs + "f").Remove();
                cell.Elements(mainNs + "v").Remove();
                cell.Elements(mainNs + "is").Remove();
                cell.Attribute("t")?.Remove();
            }
        }

        private static void InsertRowInOrder(XElement sheetDataElement, XNamespace mainNs, XElement row)
        {
            if (!int.TryParse((string?)row.Attribute("r"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex))
            {
                sheetDataElement.Add(row);
                return;
            }

            var nextRow = sheetDataElement
                .Elements(mainNs + "row")
                .FirstOrDefault(r => int.TryParse((string?)r.Attribute("r"), out var rIndex) && rIndex > rowIndex);

            if (nextRow == null)
            {
                sheetDataElement.Add(row);
                return;
            }

            nextRow.AddBeforeSelf(row);
        }

        private static string ReplaceCellRow(string cellReference, int newRowIndex)
        {
            var match = Regex.Match(cellReference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return cellReference;
            }

            return $"{match.Groups[1].Value.ToUpperInvariant()}{newRowIndex.ToString(CultureInfo.InvariantCulture)}";
        }

        private static string ShiftCellReference(string cellReference, int delta)
        {
            var match = Regex.Match(cellReference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return cellReference;
            }

            var rowIndex = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) + delta;
            return $"{match.Groups[1].Value.ToUpperInvariant()}{rowIndex.ToString(CultureInfo.InvariantCulture)}";
        }

        private static bool TryParseRange(
            string rangeReference,
            out string startColumn,
            out int startRow,
            out string endColumn,
            out int endRow)
        {
            startColumn = string.Empty;
            endColumn = string.Empty;
            startRow = 0;
            endRow = 0;

            var match = Regex.Match(
                rangeReference,
                @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$",
                RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                return false;
            }

            startColumn = match.Groups[1].Value.ToUpperInvariant();
            endColumn = match.Groups[3].Value.ToUpperInvariant();

            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out startRow))
            {
                return false;
            }

            if (!int.TryParse(match.Groups[4].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out endRow))
            {
                return false;
            }

            return true;
        }

        private static void PopulateInvoiceSheet(
            XElement sheetDataElement,
            XNamespace mainNs,
            Customer customer,
            string periodLabel,
            string issuerName,
            VisibleZeroStyleResolver? visibleZeroStyleResolver)
        {
            var period = periodLabel?.Trim() ?? string.Empty;
            var issuer = issuerName?.Trim() ?? string.Empty;

            var householdName = customer.Name?.Trim() ?? string.Empty;
            var receiptNumber = customer.SequenceNumber > 0 ? customer.SequenceNumber : 1;

            var representativeName = customer.RepresentativeName?.Trim() ?? string.Empty;
            var groupName = TextNormalization.NormalizeForDisplay(customer.GroupName);
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
            var hasValidReading = !customer.IsMissingReading && !customer.HasReadingError;

            UpdateNumberCell(sheetDataElement, mainNs, "B13", customer.CurrentIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "C13", customer.PreviousIndex);
            UpdateNumberCell(sheetDataElement, mainNs, "D13", multiplier);
            UpdateNumberCell(sheetDataElement, mainNs, "F13", customer.SubsidizedKwh);
            UpdateNumberCell(sheetDataElement, mainNs, "G13", customer.UnitPrice);

            if (hasValidReading)
            {
                UpdateNumberCellVisibleZero(sheetDataElement, sheetDataElement.Parent as XElement, mainNs, "E13", consumption, visibleZeroStyleResolver);
                UpdateNumberCellVisibleZero(sheetDataElement, sheetDataElement.Parent as XElement, mainNs, "H13", amount, visibleZeroStyleResolver);
                UpdateNumberCellVisibleZero(sheetDataElement, sheetDataElement.Parent as XElement, mainNs, "H18", amount, visibleZeroStyleResolver);
            }
            else
            {
                UpdateNumberCell(sheetDataElement, mainNs, "E13", null);
                UpdateNumberCell(sheetDataElement, mainNs, "H13", null);
                UpdateNumberCell(sheetDataElement, mainNs, "H18", null);
            }

            if (!string.IsNullOrWhiteSpace(location))
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", $"Vị trí đặt: {location}.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "I13", string.Empty);
            }

            if (hasValidReading)
            {
                var amountText = VietnameseNumberTextService.ConvertAmountToText(amount);
                UpdateTextCell(
                    sheetDataElement,
                    mainNs,
                    "A19",
                    string.IsNullOrWhiteSpace(amountText) ? string.Empty : $"Bằng chữ: {amountText}./.");
            }
            else
            {
                UpdateTextCell(sheetDataElement, mainNs, "A19", string.Empty);
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

        private static void UpdateNumberCellVisibleZero(
            XElement sheetDataElement,
            XElement? worksheetRoot,
            XNamespace ns,
            string cellReference,
            decimal? value,
            VisibleZeroStyleResolver? visibleZeroStyleResolver)
        {
            if (value is null)
            {
                UpdateNumberCell(sheetDataElement, ns, cellReference, null);
                return;
            }

            UpdateNumberCell(sheetDataElement, ns, cellReference, value);

            if (value.Value != 0m || visibleZeroStyleResolver == null)
            {
                return;
            }

            var cell = GetCell(sheetDataElement, ns, cellReference, createIfMissing: true);
            if (cell == null)
            {
                return;
            }

            var baseStyle = (string?)cell.Attribute("s");
            if (string.IsNullOrWhiteSpace(baseStyle) && worksheetRoot != null)
            {
                var colIndex = TryGetColumnIndex(cellReference);
                if (colIndex > 0)
                {
                    baseStyle = TryGetColumnStyleIndex(worksheetRoot, ns, colIndex);
                }
            }

            var visibleZeroStyle = visibleZeroStyleResolver.GetVisibleZeroStyleIndex(baseStyle);
            if (!string.IsNullOrWhiteSpace(visibleZeroStyle))
            {
                cell.SetAttributeValue("s", visibleZeroStyle);
            }
        }

        private static void ApplyShrinkToFitTextStyle(
            XElement sheetDataElement,
            XElement? worksheetRoot,
            XNamespace ns,
            string cellReference,
            VisibleZeroStyleResolver? styleResolver)
        {
            if (styleResolver == null)
            {
                return;
            }

            var cell = GetCell(sheetDataElement, ns, cellReference, createIfMissing: false);
            if (cell == null)
            {
                return;
            }

            var baseStyle = (string?)cell.Attribute("s");
            if (string.IsNullOrWhiteSpace(baseStyle) && worksheetRoot != null)
            {
                var colIndex = TryGetColumnIndex(cellReference);
                if (colIndex > 0)
                {
                    baseStyle = TryGetColumnStyleIndex(worksheetRoot, ns, colIndex);
                }
            }
            var shrinkStyle = styleResolver.GetShrinkToFitStyleIndex(baseStyle);
            if (!string.IsNullOrWhiteSpace(shrinkStyle))
            {
                cell.SetAttributeValue("s", shrinkStyle);
            }
        }

        private static int TryGetColumnIndex(string cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference))
            {
                return 0;
            }

            var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
            if (letters.Length == 0)
            {
                return 0;
            }

            var col = 0;
            foreach (var ch in letters)
            {
                if (ch < 'A' || ch > 'Z')
                {
                    return 0;
                }

                col = (col * 26) + (ch - 'A' + 1);
            }

            return col;
        }

        private static string? TryGetColumnStyleIndex(XElement worksheetRoot, XNamespace ns, int columnIndex)
        {
            if (columnIndex <= 0)
            {
                return null;
            }

            var cols = worksheetRoot.Element(ns + "cols");
            if (cols == null)
            {
                return null;
            }

            foreach (var col in cols.Elements(ns + "col"))
            {
                var min = (int?)col.Attribute("min") ?? 0;
                var max = (int?)col.Attribute("max") ?? 0;
                if (min <= 0 || max <= 0)
                {
                    continue;
                }

                if (columnIndex < min || columnIndex > max)
                {
                    continue;
                }

                var style = (string?)col.Attribute("style");
                return string.IsNullOrWhiteSpace(style) ? null : style;
            }

            return null;
        }

        private sealed class VisibleZeroStyleResolver
        {
            private const string StylesEntryPath = "xl/styles.xml";
            private const int DefaultCustomNumFmtStart = 164;

            private readonly ZipArchiveEntry stylesEntry;
            private readonly XDocument stylesDoc;
            private readonly XNamespace ns;
            private readonly Dictionary<int, int> visibleZeroStyleCache = new();
            private readonly Dictionary<int, int> shrinkToFitStyleCache = new();

            private bool isDirty;

            private VisibleZeroStyleResolver(ZipArchiveEntry stylesEntry, XDocument stylesDoc)
            {
                this.stylesEntry = stylesEntry;
                this.stylesDoc = stylesDoc;
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            }

            public static VisibleZeroStyleResolver? TryLoad(ZipArchive archive)
            {
                var entry = archive.GetEntry(StylesEntryPath);
                if (entry == null)
                {
                    return null;
                }

                XDocument doc;
                using (var stream = entry.Open())
                {
                    doc = XDocument.Load(stream);
                }

                return new VisibleZeroStyleResolver(entry, doc);
            }

            public string? GetVisibleZeroStyleIndex(string? baseStyleIndex)
            {
                if (string.IsNullOrWhiteSpace(baseStyleIndex))
                {
                    return baseStyleIndex;
                }

                if (!int.TryParse(baseStyleIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out var baseIndex) ||
                    baseIndex < 0)
                {
                    return baseStyleIndex;
                }

                if (visibleZeroStyleCache.TryGetValue(baseIndex, out var cached))
                {
                    return cached.ToString(CultureInfo.InvariantCulture);
                }

                var styleSheet = stylesDoc.Root;
                if (styleSheet == null)
                {
                    return baseStyleIndex;
                }

                var cellXfs = styleSheet.Element(ns + "cellXfs");
                if (cellXfs == null)
                {
                    return baseStyleIndex;
                }

                var xfs = cellXfs.Elements(ns + "xf").ToList();
                if (baseIndex >= xfs.Count)
                {
                    return baseStyleIndex;
                }

                var baseXf = xfs[baseIndex];
                var numFmtId = (int?)baseXf.Attribute("numFmtId") ?? 0;
                if (!TryGetFormatCode(styleSheet, numFmtId, out var formatCode))
                {
                    return baseStyleIndex;
                }

                if (!TryMakeZeroVisibleFormat(formatCode, out var visibleZeroFormat))
                {
                    return baseStyleIndex;
                }

                var newNumFmtId = GetOrCreateNumFmtId(styleSheet, visibleZeroFormat);
                var newXf = new XElement(baseXf);
                newXf.SetAttributeValue("numFmtId", newNumFmtId.ToString(CultureInfo.InvariantCulture));
                newXf.SetAttributeValue("applyNumberFormat", "1");
                cellXfs.Add(newXf);

                var newStyleIndex = xfs.Count;
                cellXfs.SetAttributeValue("count", (newStyleIndex + 1).ToString(CultureInfo.InvariantCulture));

                visibleZeroStyleCache[baseIndex] = newStyleIndex;
                isDirty = true;

                return newStyleIndex.ToString(CultureInfo.InvariantCulture);
            }

            public string? GetShrinkToFitStyleIndex(string? baseStyleIndex)
            {
                if (string.IsNullOrWhiteSpace(baseStyleIndex))
                {
                    return baseStyleIndex;
                }

                if (!int.TryParse(baseStyleIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out var baseIndex) ||
                    baseIndex < 0)
                {
                    return baseStyleIndex;
                }

                if (shrinkToFitStyleCache.TryGetValue(baseIndex, out var cached))
                {
                    return cached.ToString(CultureInfo.InvariantCulture);
                }

                var styleSheet = stylesDoc.Root;
                if (styleSheet == null)
                {
                    return baseStyleIndex;
                }

                var cellXfs = styleSheet.Element(ns + "cellXfs");
                if (cellXfs == null)
                {
                    return baseStyleIndex;
                }

                var xfs = cellXfs.Elements(ns + "xf").ToList();
                if (baseIndex >= xfs.Count)
                {
                    return baseStyleIndex;
                }

                var baseXf = xfs[baseIndex];
                var baseAlignment = baseXf.Element(ns + "alignment");
                if (baseAlignment != null && string.Equals((string?)baseAlignment.Attribute("shrinkToFit"), "1", StringComparison.Ordinal))
                {
                    shrinkToFitStyleCache[baseIndex] = baseIndex;
                    return baseStyleIndex;
                }

                var newXf = new XElement(baseXf);
                var alignment = newXf.Element(ns + "alignment");
                if (alignment == null)
                {
                    alignment = new XElement(ns + "alignment");
                    newXf.Add(alignment);
                }

                alignment.SetAttributeValue("shrinkToFit", "1");
                if (string.IsNullOrWhiteSpace((string?)alignment.Attribute("horizontal")))
                {
                    alignment.SetAttributeValue("horizontal", "left");
                }

                newXf.SetAttributeValue("applyAlignment", "1");
                cellXfs.Add(newXf);

                var newStyleIndex = xfs.Count;
                cellXfs.SetAttributeValue("count", (newStyleIndex + 1).ToString(CultureInfo.InvariantCulture));

                shrinkToFitStyleCache[baseIndex] = newStyleIndex;
                isDirty = true;

                return newStyleIndex.ToString(CultureInfo.InvariantCulture);
            }

            public void SaveIfDirty()
            {
                if (!isDirty)
                {
                    return;
                }

                using var stream = stylesEntry.Open();
                stream.SetLength(0);
                stylesDoc.Save(stream);
            }

            private static bool TryGetFormatCode(XElement styleSheet, int numFmtId, out string formatCode)
            {
                formatCode = string.Empty;

                var ns = styleSheet.Name.Namespace;
                var numFmts = styleSheet.Element(ns + "numFmts");
                if (numFmts == null)
                {
                    return false;
                }

                var numFmt = numFmts
                    .Elements(ns + "numFmt")
                    .FirstOrDefault(e => (int?)e.Attribute("numFmtId") == numFmtId);

                formatCode = (string?)numFmt?.Attribute("formatCode") ?? string.Empty;
                return !string.IsNullOrWhiteSpace(formatCode);
            }

            private int GetOrCreateNumFmtId(XElement styleSheet, string formatCode)
            {
                var numFmts = styleSheet.Element(ns + "numFmts");
                if (numFmts == null)
                {
                    numFmts = new XElement(ns + "numFmts");
                    var cellStyleXfs = styleSheet.Element(ns + "cellStyleXfs");
                    if (cellStyleXfs != null)
                    {
                        cellStyleXfs.AddBeforeSelf(numFmts);
                    }
                    else
                    {
                        styleSheet.AddFirst(numFmts);
                    }
                }

                var existing = numFmts
                    .Elements(ns + "numFmt")
                    .FirstOrDefault(e => string.Equals((string?)e.Attribute("formatCode"), formatCode, StringComparison.Ordinal));

                if (existing != null)
                {
                    return (int?)existing.Attribute("numFmtId") ?? DefaultCustomNumFmtStart;
                }

                var maxId = numFmts
                    .Elements(ns + "numFmt")
                    .Select(e => (int?)e.Attribute("numFmtId"))
                    .Where(v => v != null)
                    .Select(v => v!.Value)
                    .DefaultIfEmpty(DefaultCustomNumFmtStart - 1)
                    .Max();

                var newId = Math.Max(DefaultCustomNumFmtStart, maxId + 1);
                numFmts.Add(new XElement(ns + "numFmt",
                    new XAttribute("numFmtId", newId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("formatCode", formatCode)));

                var count = numFmts.Elements(ns + "numFmt").Count();
                numFmts.SetAttributeValue("count", count.ToString(CultureInfo.InvariantCulture));

                isDirty = true;
                return newId;
            }

            private static bool TryMakeZeroVisibleFormat(string formatCode, out string visibleZeroFormat)
            {
                visibleZeroFormat = string.Empty;
                if (string.IsNullOrWhiteSpace(formatCode))
                {
                    return false;
                }

                var sections = SplitFormatSections(formatCode);
                if (sections.Count < 3)
                {
                    return false;
                }

                if (string.Equals(sections[2], sections[0], StringComparison.Ordinal))
                {
                    return false;
                }

                sections[2] = sections[0];
                visibleZeroFormat = string.Join(";", sections);
                return !string.IsNullOrWhiteSpace(visibleZeroFormat) &&
                       !string.Equals(visibleZeroFormat, formatCode, StringComparison.Ordinal);
            }

            private static List<string> SplitFormatSections(string formatCode)
            {
                var sections = new List<string>();
                var current = new StringBuilder();
                var inQuotes = false;

                foreach (var ch in formatCode)
                {
                    if (ch == '"')
                    {
                        inQuotes = !inQuotes;
                        current.Append(ch);
                        continue;
                    }

                    if (ch == ';' && !inQuotes)
                    {
                        sections.Add(current.ToString());
                        current.Clear();
                        continue;
                    }

                    current.Append(ch);
                }

                sections.Add(current.ToString());
                return sections;
            }
        }

        private static void CopyCellStyle(
            XElement sheetDataElement,
            XNamespace ns,
            string sourceCellReference,
            string targetCellReference)
        {
            if (string.IsNullOrWhiteSpace(sourceCellReference) || string.IsNullOrWhiteSpace(targetCellReference))
            {
                return;
            }

            var sourceCell = GetCell(sheetDataElement, ns, sourceCellReference, createIfMissing: false);
            var styleAttr = (string?)sourceCell?.Attribute("s");
            if (string.IsNullOrWhiteSpace(styleAttr))
            {
                return;
            }

            var targetCell = GetCell(sheetDataElement, ns, targetCellReference, createIfMissing: true);
            if (targetCell == null)
            {
                return;
            }

            targetCell.SetAttributeValue("s", styleAttr);
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

        private static XElement? GetCell(
            XElement sheetDataElement,
            XNamespace ns,
            string cellReference,
            bool createIfMissing)
        {
            var rowIndex = GetRowIndex(cellReference);
            if (rowIndex <= 0)
            {
                return null;
            }

            var row = sheetDataElement
                .Elements(ns + "row")
                .FirstOrDefault(r => string.Equals(
                    (string?)r.Attribute("r"),
                    rowIndex.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal));

            if (row == null)
            {
                return null;
            }

            var cell = row
                .Elements(ns + "c")
                .FirstOrDefault(c => string.Equals(
                    (string?)c.Attribute("r"),
                    cellReference,
                    StringComparison.OrdinalIgnoreCase));

            if (cell == null && createIfMissing)
            {
                cell = new XElement(ns + "c", new XAttribute("r", cellReference));
                row.Add(cell);
            }

            return cell;
        }

        private static void EnsureWorksheetShowsZeros(XElement worksheetRoot, XNamespace mainNs)
        {
            var sheetViews = worksheetRoot.Element(mainNs + "sheetViews");
            if (sheetViews == null)
            {
                return;
            }

            foreach (var view in sheetViews.Elements(mainNs + "sheetView"))
            {
                view.SetAttributeValue("showZeros", "1");
            }
        }

        private static void EnsureInvoiceColumnWidths(XElement worksheetRoot, XNamespace mainNs)
        {
            // Excel shows #### when a numeric value doesn't fit the column width.
            // Amount columns (H/I) must handle big totals (hundreds of millions+).
            const double amountColumnWidth = 24.7109375;

            EnsureMinColumnWidth(worksheetRoot, mainNs, 8, amountColumnWidth); // H
            EnsureMinColumnWidth(worksheetRoot, mainNs, 9, amountColumnWidth); // I
        }

        private static void EnsureMinColumnWidth(
            XElement worksheetRoot,
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
                var sheetData = worksheetRoot.Element(mainNs + "sheetData");
                if (sheetData != null)
                {
                    sheetData.AddBeforeSelf(cols);
                }
                else
                {
                    worksheetRoot.Add(cols);
                }
            }

            XElement? bestMatch = null;
            var bestSpan = int.MaxValue;

            foreach (var col in cols.Elements(mainNs + "col"))
            {
                var min = (int?)col.Attribute("min") ?? 0;
                var max = (int?)col.Attribute("max") ?? 0;
                if (min <= 0 || max <= 0 || columnIndex < min || columnIndex > max)
                {
                    continue;
                }

                var span = max - min;
                if (span < bestSpan)
                {
                    bestSpan = span;
                    bestMatch = col;
                }
            }

            if (bestMatch == null)
            {
                var newCol = new XElement(mainNs + "col",
                    new XAttribute("min", columnIndex),
                    new XAttribute("max", columnIndex),
                    new XAttribute("width", minWidth.ToString("0.########", CultureInfo.InvariantCulture)),
                    new XAttribute("customWidth", 1));

                var inserted = false;
                foreach (var existing in cols.Elements(mainNs + "col").ToList())
                {
                    var existingMin = (int?)existing.Attribute("min") ?? int.MaxValue;
                    if (existingMin > columnIndex)
                    {
                        existing.AddBeforeSelf(newCol);
                        inserted = true;
                        break;
                    }
                }

                if (!inserted)
                {
                    cols.Add(newCol);
                }

                return;
            }

            var widthText = (string?)bestMatch.Attribute("width");
            var currentWidth = double.TryParse(widthText, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : 0;

            if (currentWidth < minWidth)
            {
                bestMatch.SetAttributeValue("width", minWidth.ToString("0.########", CultureInfo.InvariantCulture));
                bestMatch.SetAttributeValue("customWidth", 1);
            }
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
