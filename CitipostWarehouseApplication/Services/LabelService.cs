using CitipostWarehouseApplication.Models;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace CitipostWarehouseApplication.Services
{
    public class LabelService
    {
        private List<LabelData> labelsToPrint;
        private int currentLabelIndex = 0;

        public void GenerateLabels(List<Subscriber> subscribers)
        {
            labelsToPrint = subscribers.Select(s => new LabelData
            {
                // Map properties from Subscriber to LabelData
                ContactName = s.ContactFullName,
                AccountName = s.AccountName,
                FullAddress = FormatAddress(s),
                SupplierName = s.AssignedSupplier,
                PrintDate = DateTime.Now
            }).ToList();
        }

        private string FormatAddress(Subscriber subscriber)
        {
            var addressParts = new List<string>();

            if (!string.IsNullOrEmpty(subscriber.Address1))
                addressParts.Add(subscriber.Address1);
            if (!string.IsNullOrEmpty(subscriber.Address2))
                addressParts.Add(subscriber.Address2);
            if (!string.IsNullOrEmpty(subscriber.Address3))
                addressParts.Add(subscriber.Address3);
            if (!string.IsNullOrEmpty(subscriber.City))
                addressParts.Add(subscriber.City);
            if (!string.IsNullOrEmpty(subscriber.StateProvince))
                addressParts.Add(subscriber.StateProvince);
            if (!string.IsNullOrEmpty(subscriber.PostCode))
                addressParts.Add(subscriber.PostCode);
            if (!string.IsNullOrEmpty(subscriber.Country))
                addressParts.Add(subscriber.Country);

            return string.Join(", ", addressParts);
        }

        public void PrintLabels()
        {
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrintPage += PrintDoc_PrintPage;

            currentLabelIndex = 0; // Reset index for printing

            // Configure for label printer (adjust as needed)
            printDoc.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);
            printDoc.DefaultPageSettings.Margins = new Margins(30, 50, 50, 50); // Adjust margins as needed

            try
            {
                printDoc.Print();
            }
            catch (Exception ex)
            {
                throw new Exception($"Printing failed: {ex.Message}");
            }
        }

        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;

            // Label dimensions and layout
            float labelWidth = 350;   
            float labelHeight = 150;  
            int labelsPerRow = 2;     
            int labelsPerColumn = 5;  
            int labelsPerPage = labelsPerRow * labelsPerColumn; // Total labels per page

            // Define margins
            float leftMargin = 50;
            float topMargin = 50;
            float horizontalSpacing = 20; // Space between labels horizontally
            float verticalSpacing = 30;   // Space between labels vertically

            // Define fonts
            Font headerFont = new Font("Arial", 11, FontStyle.Bold);
            Font bodyFont = new Font("Arial", 9);
            Font smallFont = new Font("Arial", 7);

            // Calculate how many labels to print on this page
            int labelsOnThisPage = Math.Min(labelsPerPage, labelsToPrint.Count - currentLabelIndex);

            for (int i = 0; i < labelsOnThisPage; i++)
            {
                if (currentLabelIndex >= labelsToPrint.Count)
                    break;

                var label = labelsToPrint[currentLabelIndex];

                // Calculate position for this label
                int row = i / labelsPerRow;
                int col = i % labelsPerRow;

                // Calculate x and y position based on row and column
                float x = leftMargin + (col * (labelWidth + horizontalSpacing));
                float y = topMargin + (row * (labelHeight + verticalSpacing));

                // Draw label content
                float currentY = y + 8;
                float lineHeight = 13;

                // Draw header
                g.DrawString($"TO: {label.ContactName}", headerFont, Brushes.Black, x + 8, currentY);
                currentY += lineHeight + 3;

                g.DrawString($"Account: {label.AccountName}", bodyFont, Brushes.Black, x + 8, currentY);
                currentY += lineHeight;

                // Handle long addresses by wrapping text
                string[] addressLines = WrapText(g, $"Address: {label.FullAddress}", bodyFont, labelWidth - 16);
                foreach (string line in addressLines)
                {
                    g.DrawString(line, bodyFont, Brushes.Black, x + 10, currentY);
                    currentY += lineHeight;
                }
                currentY += lineHeight;

                g.DrawString($"Supplier: {label.SupplierName}", bodyFont, Brushes.Black, x + 8, currentY);
                currentY += lineHeight;

                g.DrawString($"Date: {label.PrintDate:yyyy-MM-dd HH:mm}", smallFont, Brushes.Black, x + 8, currentY);

                // Draw border around label
                g.DrawRectangle(Pens.Black, x, y, labelWidth, labelHeight);

                currentLabelIndex++;
            }

            // Check if there are more pages to print
            e.HasMorePages = currentLabelIndex < labelsToPrint.Count;
        }

        // Helper method to wrap text if it's too long
        private string[] WrapText(Graphics g, string text, Font font, float maxWidth)
        {
            List<string> lines = new List<string>();
            string[] words = text.Split(' ');
            string currentLine = "";

            foreach (string word in words)
            {
                // Check if adding the next word exceeds the maximum width
                string testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                SizeF size = g.MeasureString(testLine, font);

                if (size.Width > maxWidth && !string.IsNullOrEmpty(currentLine))
                {
                    lines.Add(currentLine);
                    currentLine = word;
                }
                else
                {
                    currentLine = testLine;
                }
            }

            if (!string.IsNullOrEmpty(currentLine))
                lines.Add(currentLine);

            return lines.ToArray();
        }

        public void ExportLabelsToExcel(string filePath)
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Carrier Labels");

                // Headers
                worksheet.Cells[1, 1].Value = "Contact Name";
                worksheet.Cells[1, 2].Value = "Account Name";
                worksheet.Cells[1, 3].Value = "Full Address";
                worksheet.Cells[1, 4].Value = "Supplier";
                worksheet.Cells[1, 5].Value = "Print Date";

                // Data
                for (int i = 0; i < labelsToPrint.Count; i++)
                {
                    var label = labelsToPrint[i];
                    int row = i + 2;

                    worksheet.Cells[row, 1].Value = label.ContactName;
                    worksheet.Cells[row, 2].Value = label.AccountName;
                    worksheet.Cells[row, 3].Value = label.FullAddress;
                    worksheet.Cells[row, 4].Value = label.SupplierName;
                    worksheet.Cells[row, 5].Value = label.PrintDate;
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                package.SaveAs(new System.IO.FileInfo(filePath));
            }
        }

        public int GetLabelCount()
        {
            return labelsToPrint?.Count ?? 0;
        }
    }
}
