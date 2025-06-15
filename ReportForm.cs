using CitipostWarehouseApplication.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CitipostWarehouseApplication
{
    public partial class ReportForm : Form
    {
        private List<SummaryReport> summaryData;
        private DataGridView reportGrid;
        private Button btnExportReport, btnClose;
        private Label lblTotalSuppliers, lblTotalItems;

        public ReportForm(List<SummaryReport> summaryData)
        {
            this.summaryData = summaryData;
            InitializeComponent();
            LoadReportData();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(800, 600);
            this.Text = "Supplier Summary Report";
            this.StartPosition = FormStartPosition.CenterParent;

            CreateControls();
            SetupLayout();
            WireEvents();
        }

        private void CreateControls()
        {
            // Summary labels
            lblTotalSuppliers = new Label
            {
                Text = $"Total Suppliers: {summaryData.Count}",
                Size = new Size(200, 20),
                Font = new Font("Arial", 10, FontStyle.Bold)
            };

            lblTotalItems = new Label
            {
                Text = $"Total Items: {summaryData.Sum(s => s.TotalItems)}",
                Size = new Size(200, 20),
                Font = new Font("Arial", 10, FontStyle.Bold)
            };

            // Report grid
            reportGrid = new DataGridView
            {
                Size = new Size(760, 450),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            // Buttons
            btnExportReport = new Button { Text = "Export to Excel", Size = new Size(120, 30) };
            btnClose = new Button { Text = "Close", Size = new Size(80, 30) };
        }

        private void SetupLayout()
        {
            // Title
            Label titleLabel = new Label
            {
                Text = "Warehouse Operations Summary Report",
                Font = new Font("Arial", 14, FontStyle.Bold),
                Size = new Size(400, 25),
                Location = new Point(20, 20)
            };

            // Summary section
            Label summaryTitle = new Label
            {
                Text = "Summary Information:",
                Font = new Font("Arial", 12, FontStyle.Bold),
                Size = new Size(200, 20),
                Location = new Point(20, 60)
            };

            lblTotalSuppliers.Location = new Point(20, 85);
            lblTotalItems.Location = new Point(20, 110);

            // Report date
            Label reportDate = new Label
            {
                Text = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                Size = new Size(200, 20),
                Location = new Point(20, 135)
            };

            // Grid
            reportGrid.Location = new Point(20, 170);

            // Buttons
            btnExportReport.Location = new Point(600, 520);
            btnClose.Location = new Point(690, 520);

            // Add controls to form
            this.Controls.AddRange(new Control[]
            {
            titleLabel, summaryTitle, lblTotalSuppliers, lblTotalItems,
            reportDate, reportGrid, btnExportReport, btnClose
            });
        }

        private void WireEvents()
        {
            btnExportReport.Click += BtnExportReport_Click;
            btnClose.Click += BtnClose_Click;
        }

        private void LoadReportData()
        {
            // Create a flattened view for the grid
            var reportData = summaryData.Select(s => new
            {
                SupplierName = s.SupplierName,
                TotalItems = s.TotalItems,
                Countries = string.Join(", ", s.Countries),
                Percentage = Math.Round((double)s.TotalItems / summaryData.Sum(x => x.TotalItems) * 100, 2)
            }).OrderByDescending(x => x.TotalItems).ToList();

            reportGrid.DataSource = reportData;

            // Customize column headers
            if (reportGrid.Columns.Count > 0)
            {
                reportGrid.Columns[0].HeaderText = "Supplier Name";
                reportGrid.Columns[1].HeaderText = "Total Items";
                reportGrid.Columns[2].HeaderText = "Countries Served";
                reportGrid.Columns[3].HeaderText = "Percentage %";
            }
        }

        private void BtnExportReport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"SupplierSummaryReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportToExcel(saveDialog.FileName);
                    MessageBox.Show($"Report exported successfully to:\n{saveDialog.FileName}",
                        "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting report: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Supplier Summary");

                // Title
                worksheet.Cells[1, 1].Value = "Warehouse Operations Summary Report";
                worksheet.Cells[1, 1, 1, 4].Merge = true;
                worksheet.Cells[1, 1].Style.Font.Size = 16;
                worksheet.Cells[1, 1].Style.Font.Bold = true;

                // Report date
                worksheet.Cells[2, 1].Value = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                worksheet.Cells[2, 1, 2, 4].Merge = true;

                // Summary information
                worksheet.Cells[4, 1].Value = "Summary Information:";
                worksheet.Cells[4, 1].Style.Font.Bold = true;
                worksheet.Cells[5, 1].Value = $"Total Suppliers: {summaryData.Count}";
                worksheet.Cells[6, 1].Value = $"Total Items: {summaryData.Sum(s => s.TotalItems)}";

                // Headers
                int startRow = 8;
                worksheet.Cells[startRow, 1].Value = "Supplier Name";
                worksheet.Cells[startRow, 2].Value = "Total Items";
                worksheet.Cells[startRow, 3].Value = "Countries Served";
                worksheet.Cells[startRow, 4].Value = "Percentage %";

                // Make headers bold
                worksheet.Cells[startRow, 1, startRow, 4].Style.Font.Bold = true;

                // Data
                var orderedData = summaryData.OrderByDescending(s => s.TotalItems).ToList();
                for (int i = 0; i < orderedData.Count; i++)
                {
                    int row = startRow + 1 + i;
                    var supplier = orderedData[i];

                    worksheet.Cells[row, 1].Value = supplier.SupplierName;
                    worksheet.Cells[row, 2].Value = supplier.TotalItems;
                    worksheet.Cells[row, 3].Value = string.Join(", ", supplier.Countries);
                    worksheet.Cells[row, 4].Value = Math.Round((double)supplier.TotalItems / summaryData.Sum(x => x.TotalItems) * 100, 2);
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                // Add borders
                var dataRange = worksheet.Cells[startRow, 1, startRow + orderedData.Count, 4];
                dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
