using CitipostWarehouseApplication.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CitipostWarehouseApplication
{
    public partial class MainForm : Form
    {
        private DataService dataService;
        private LabelService labelService;
        private DataGridView dataGridView;
        private Button btnLoadData, btnAssignSuppliers, btnGenerateLabels, btnPrintLabels, btnGenerateReport, btnSaveData;
        private Label lblStatus;
        private TextBox txtFileStatus;

        public MainForm()
        {
            InitializeComponent();
            dataService = new DataService();
            labelService = new LabelService();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(1200, 800);
            this.Text = "Warehouse Operations Management System";
            this.StartPosition = FormStartPosition.CenterScreen;

            // Create controls
            CreateControls();
            SetupLayout();
            WireEvents();
        }

        private void CreateControls()
        {
            // Buttons
            btnLoadData = new Button { Text = "1. Load Data Files", Size = new Size(120, 30) };
            btnAssignSuppliers = new Button { Text = "2. Assign Suppliers", Size = new Size(120, 30) };
            btnGenerateLabels = new Button { Text = "3. Generate Labels", Size = new Size(120, 30) };
            btnPrintLabels = new Button { Text = "4. Print Labels", Size = new Size(120, 30) };
            btnGenerateReport = new Button { Text = "5. Generate Report", Size = new Size(120, 30) };
            btnSaveData = new Button { Text = "6. Save Data", Size = new Size(120, 30) };

            // Status controls
            lblStatus = new Label { Text = "Ready", Size = new Size(200, 20), ForeColor = Color.Green };
            txtFileStatus = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Size = new Size(300, 100)
            };

            // Data grid
            dataGridView = new DataGridView
            {
                Size = new Size(1150, 500),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
        }

        private void SetupLayout()
        {
            // Create panels for better organization
            Panel topPanel = new Panel { Size = new Size(1180, 50), Location = new Point(10, 10) };
            Panel buttonPanel = new Panel { Size = new Size(1180, 40), Location = new Point(10, 70) };
            Panel statusPanel = new Panel { Size = new Size(1180, 120), Location = new Point(10, 120) };
            Panel gridPanel = new Panel { Size = new Size(1180, 520), Location = new Point(10, 250) };

            // Top panel
            Label titleLabel = new Label
            {
                Text = "Warehouse Operations Management System",
                Font = new Font("Arial", 16, FontStyle.Bold),
                Size = new Size(500, 30),
                Location = new Point(10, 10)
            };
            topPanel.Controls.Add(titleLabel);

            // Button panel
            btnLoadData.Location = new Point(10, 5);
            btnAssignSuppliers.Location = new Point(140, 5);
            btnGenerateLabels.Location = new Point(270, 5);
            btnPrintLabels.Location = new Point(400, 5);
            btnGenerateReport.Location = new Point(530, 5);
            btnSaveData.Location = new Point(660, 5);

            buttonPanel.Controls.AddRange(new Control[]
            {
            btnLoadData, btnAssignSuppliers, btnGenerateLabels,
            btnPrintLabels, btnGenerateReport, btnSaveData
            });

            // Status panel
            Label statusLabelTitle = new Label { Text = "Status:", Location = new Point(10, 10), Size = new Size(50, 20) };
            lblStatus.Location = new Point(70, 10);

            Label fileStatusTitle = new Label { Text = "File Information:", Location = new Point(10, 40), Size = new Size(100, 20) };
            txtFileStatus.Location = new Point(10, 60);

            statusPanel.Controls.AddRange(new Control[] { statusLabelTitle, lblStatus, fileStatusTitle, txtFileStatus });

            // Grid panel
            dataGridView.Location = new Point(10, 10);
            gridPanel.Controls.Add(dataGridView);

            // Add panels to form
            this.Controls.AddRange(new Control[] { topPanel, buttonPanel, statusPanel, gridPanel });
        }

        private void WireEvents()
        {
            // Wire up button click events
            btnLoadData.Click += BtnLoadData_Click;
            btnAssignSuppliers.Click += BtnAssignSuppliers_Click;
            btnGenerateLabels.Click += BtnGenerateLabels_Click;
            btnPrintLabels.Click += BtnPrintLabels_Click;
            btnGenerateReport.Click += BtnGenerateReport_Click;
            btnSaveData.Click += BtnSaveData_Click;
        }

        private void BtnLoadData_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus("Loading data files...", Color.Orange);

                // Get project root directory
                string projectRoot = GetProjectRootDirectory();

                // Check if files exist in Data folder
                string[] requiredFiles =
                {
                    Path.Combine(projectRoot, "Data", "HFM#526-Subscribers Report.xlsx"),
                    Path.Combine(projectRoot, "Data", "PAGEANT ROUTING GUIDE.xlsx"),
                    Path.Combine(projectRoot, "Data", "INTL TEMPLATE.doc")
                };

                var missingFiles = requiredFiles.Where(f => !File.Exists(f)).ToList();

                if (missingFiles.Any())
                {
                    txtFileStatus.Text = "Missing files:\r\n" + string.Join("\r\n", missingFiles.Select(Path.GetFileName));
                    txtFileStatus.Text += "\r\n\r\nPlease ensure all files are in the Data folder of the application directory.";
                    UpdateStatus("Missing required files", Color.Red);
                    return;
                }

                // Load data
                dataService = new DataService();
                var subscribers = dataService.GetSubscribers();

                // Display in grid
                dataGridView.DataSource = subscribers;

                txtFileStatus.Text = $"Successfully loaded:\r\n";
                txtFileStatus.Text += $"- Subscribers: {subscribers.Count}\r\n";
                txtFileStatus.Text += $"- Suppliers: {dataService.GetSuppliers().Count}\r\n";
                txtFileStatus.Text += $"Files processed at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

                UpdateStatus($"Data loaded successfully ({subscribers.Count} records)", Color.Green);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error loading data", Color.Red);
                MessageBox.Show($"Error loading data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAssignSuppliers_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus("Assigning suppliers...", Color.Orange);

                // Data is already processed in DataService constructor
                dataGridView.Refresh();

                var assignedCount = dataService.GetSubscribers().Count(s => !string.IsNullOrEmpty(s.AssignedSupplier));
                UpdateStatus($"Suppliers assigned ({assignedCount} assignments)", Color.Green);

                txtFileStatus.Text += $"\r\nSupplier assignment completed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            }
            catch (Exception ex)
            {
                UpdateStatus("Error assigning suppliers", Color.Red);
                MessageBox.Show($"Error assigning suppliers: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnGenerateLabels_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus("Generating labels...", Color.Orange);

                var subscribers = dataService.GetSubscribers();
                labelService.GenerateLabels(subscribers);

                // Option to export labels to Excel
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"CarrierLabels_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    labelService.ExportLabelsToExcel(saveDialog.FileName);
                    txtFileStatus.Text += $"\r\nLabels exported to: {saveDialog.FileName}";
                }

                UpdateStatus($"Labels generated ({labelService.GetLabelCount()} labels)", Color.Green);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error generating labels", Color.Red);
                MessageBox.Show($"Error generating labels: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPrintLabels_Click(object sender, EventArgs e)
        {
            try
            {
                if (labelService.GetLabelCount() == 0)
                {
                    MessageBox.Show("No labels to print. Please generate labels first.", "No Labels", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                UpdateStatus("Printing labels...", Color.Orange);

                var result = MessageBox.Show($"Print {labelService.GetLabelCount()} labels?", "Confirm Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    labelService.PrintLabels();
                    UpdateStatus("Labels printed successfully", Color.Green);
                    txtFileStatus.Text += $"\r\nLabels printed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Error printing labels", Color.Red);
                MessageBox.Show($"Error printing labels: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnGenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus("Generating summary report...", Color.Orange);

                var summaryReport = dataService.GenerateSummaryReport();

                // Create report form
                var reportForm = new ReportForm(summaryReport);
                reportForm.Show();

                UpdateStatus("Summary report generated", Color.Green);
                txtFileStatus.Text += $"\r\nSummary report generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            }
            catch (Exception ex)
            {
                UpdateStatus("Error generating report", Color.Red);
                MessageBox.Show($"Error generating report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSaveData_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateStatus("Saving processed data...", Color.Orange);

                dataService.SaveProcessedData();

                UpdateStatus("Data saved successfully", Color.Green);
                txtFileStatus.Text += $"\r\nProcessed data saved: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

                MessageBox.Show("Processed data has been saved to Excel file.", "Save Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error saving data", Color.Red);
                MessageBox.Show($"Error saving data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetProjectRootDirectory()
        {
            // Get the current executable directory
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;

            // Navigate up from bin/Debug/net8.0-windows to the project root
            DirectoryInfo dir = new DirectoryInfo(currentDir);

            // Go up until we find the project root (where .csproj file is located)
            while (dir != null && !dir.GetFiles("*.csproj").Any())
            {
                dir = dir.Parent;
            }

            if (dir == null)
            {
                throw new DirectoryNotFoundException("Could not find project root directory");
            }

            return dir.FullName;
        }

        private void UpdateStatus(string message, Color color)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = color;
            Application.DoEvents();
        }
    }
}