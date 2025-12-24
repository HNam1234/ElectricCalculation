using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ElectricCalculation.Services
{
    public static class ExcelPdfExportService
    {
        public static void ExportWorkbookToPdf(string workbookPath, string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new ArgumentException("Workbook path is required.", nameof(workbookPath));
            }

            if (!File.Exists(workbookPath))
            {
                throw new FileNotFoundException("Workbook not found.", workbookPath);
            }

            if (string.IsNullOrWhiteSpace(pdfPath))
            {
                throw new ArgumentException("PDF path is required.", nameof(pdfPath));
            }

            var directory = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Microsoft Excel is not available to export PDF.");
            }

            dynamic? excel = null;
            dynamic? workbooks = null;
            dynamic? workbook = null;

            try
            {
                excel = Activator.CreateInstance(excelType);
                if (excel == null)
                {
                    throw new InvalidOperationException("Failed to start Microsoft Excel.");
                }

                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbooks = excel.Workbooks;
                workbook = workbooks.Open(workbookPath, ReadOnly: true);

                const int xlTypePdf = 0;
                workbook.ExportAsFixedFormat(xlTypePdf, pdfPath);
            }
            finally
            {
                try
                {
                    workbook?.Close(SaveChanges: false);
                }
                catch
                {
                    // Ignore.
                }

                try
                {
                    excel?.Quit();
                }
                catch
                {
                    // Ignore.
                }

                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);
                ReleaseComObject(excel);
            }
        }

        private static void ReleaseComObject(object? comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch
            {
                // Ignore.
            }
        }
    }
}

