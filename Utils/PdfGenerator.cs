using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Proj.Utils
{
    public static class PdfGenerator
    {
        // Main method to generate PDF from image and table data
        public static byte[] Generate(Stream imageStream, List<List<string>> tableData, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            // Validate table data
            if (tableData == null || tableData.Count < 2)
                throw new ArgumentException("Table data must have at least one header row and one data row.");

            // Save image from stream to a temporary file
            var imagePath = Path.GetTempFileName();
            File.WriteAllBytes(imagePath, ReadFully(imageStream));

            var headerRow = tableData[0]; // First row is the table header
            var dataRows = tableData.Skip(1).ToList(); // Remaining rows are table data
            var document = new Document(); // Create a new PDF document

            // === PAGE 1: IMAGE PAGE ===
            var imageSection = document.AddSection();
            imageSection.PageSetup.PageWidth = Unit.FromCentimeter(21);
            imageSection.PageSetup.PageHeight = Unit.FromCentimeter(29.7 * 3 / 4); // Shorter height for image page

            // Add custom header and footer
            PdfHeaderLayout.BuildHeader(imageSection, headerModel);
            PdfFooterLayout.BuildFooter(footerModel, imageSection);

            // Add image to the center of the page
            var imageParagraph = imageSection.AddParagraph();
            imageParagraph.Format.SpaceBefore = "2cm";
            imageParagraph.Format.Alignment = ParagraphAlignment.Center;

            var image = imageParagraph.AddImage(imagePath);
            image.Width = "15cm";
            image.LockAspectRatio = true;

            // === PAGE 2+: TABLE PAGE ===
            var tableSection = document.AddSection();

            // Estimate column widths based on text length
            int columnCount = headerRow.Count;
            double charWidthCm = 0.2;     // Width per character
            double minColWidthCm = 2;     // Minimum column width
            double maxColWidthCm = 6;     // Maximum column width

            var columnWidthsCm = new double[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                int maxLen = headerRow[i].Length;
                foreach (var row in dataRows)
                {
                    if (i < row.Count && row[i] != null)
                        maxLen = Math.Max(maxLen, row[i].Length);
                }
                // Clamp width between min and max
                columnWidthsCm[i] = Math.Min(Math.Max(maxLen * charWidthCm, minColWidthCm), maxColWidthCm);
            }

            // Total table width (plus margin padding)
            double tableWidthCm = columnWidthsCm.Sum();
            double totalPageWidth = Math.Min(tableWidthCm + 3.0, 70.0);

            // Configure page size and margins
            tableSection.PageSetup.PageWidth = Unit.FromCentimeter(totalPageWidth);
            tableSection.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
            tableSection.PageSetup.RightMargin = Unit.FromCentimeter(1.5);
            tableSection.PageSetup.PageHeight = Unit.FromCentimeter(29.7);

            // Add header/footer to table section
            PdfHeaderLayout.BuildHeader(tableSection, headerModel);
            PdfFooterLayout.BuildFooter(footerModel, tableSection);

            // === BUILD TABLE ===
            var table = tableSection.AddTable();
            table.Borders.Width = 0; // Remove global borders
            table.Format.Font.Size = 9;

            // Define each column
            for (int i = 0; i < columnCount; i++)
            {
                var col = table.AddColumn(Unit.FromCentimeter(columnWidthsCm[i]));
                col.Format.Font.Size = 9;
            }

            // === HEADER ROW ===
            var headerRowObj = table.AddRow();
            headerRowObj.HeadingFormat = true;
            headerRowObj.Format.Font.Bold = true;
            headerRowObj.Format.Alignment = ParagraphAlignment.Left;

            for (int i = 0; i < columnCount; i++)
            {
                var cell = headerRowObj.Cells[i];
                var para = cell.AddParagraph(InsertSoftBreaks(headerRow[i]));
                para.Format.Alignment = ParagraphAlignment.Left;
                para.Format.Font.Size = 10; // Slightly larger for header
                cell.VerticalAlignment = VerticalAlignment.Center;
                para.Format.SpaceBefore = "0.3cm";
                para.Format.SpaceAfter = "0.3cm";

                // Only bottom border to separate header
                cell.Borders.Bottom.Width = 0.5;
            }

            // === DATA ROWS ===
            foreach (var row in dataRows)
            {
                var dataRow = table.AddRow();
                for (int i = 0; i < columnCount; i++)
                {
                    var cell = dataRow.Cells[i];
                    var text = i < row.Count ? InsertSoftBreaks(row[i]) : "";
                    var para = cell.AddParagraph(text);
                    para.Format.Alignment = ParagraphAlignment.Left;
                    para.Format.Font.Size = 9;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    para.Format.SpaceBefore = "0.1cm";
                    para.Format.SpaceAfter = "0.1cm";

                    // Add bottom border to each row
                    cell.Borders.Bottom.Width = 0.5;
                    cell.Borders.Bottom.Color = Colors.Gray;
                }
            }

            // === RENDER FINAL PDF ===
            var pdfRenderer = new PdfDocumentRenderer(true)
            {
                Document = document
            };
            pdfRenderer.RenderDocument();

            // Save PDF to memory stream and return as byte array
            using var ms = new MemoryStream();
            pdfRenderer.PdfDocument.Save(ms, false);
            return ms.ToArray();
        }

        // Helper to read full stream as byte array (for image)
        private static byte[] ReadFully(Stream input)
        {
            using var ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }

        // Inserts invisible breaks every N characters to help text wrap properly
        private static string InsertSoftBreaks(string input, int interval = 20)
        {
            if (string.IsNullOrEmpty(input) || input.Length < interval)
                return input;

            return string.Concat(input.Select((c, i) => (i > 0 && i % interval == 0) ? "\u200B" + c : c.ToString()));
        }
    }
}
