using CitipostWarehouseApplication.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CitipostWarehouseApplication.Services
{
    public class DataService
    {
        private List<Subscriber> subscribers;
        private List<Supplier> suppliers;

        public DataService()
        {
            LoadData();
        }

        private void LoadData()
        {
            LoadSubscribers();
            LoadSuppliers();
            AssignSuppliersToSubscribers();
        }

        private void LoadSubscribers()
        {
            subscribers = new List<Subscriber>();

            // Get the project root directory (go up from bin/Debug/net8.0-windows to project root)
            string projectRoot = GetProjectRootDirectory();
            string filePath = Path.Combine(projectRoot, "Data", "HFM#526-Subscribers Report.xlsx");

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The subscriber data file was not found at: {filePath}");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context for EPPlus

            // Load from HFM#526-Subscribers Report.xlsx
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new InvalidOperationException("The subscriber data file is empty or does not contain valid data.");

                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                if (worksheet.Dimension == null)
                    throw new Exception("The worksheet is empty or has no data.");

                for (int row = 2; row <= rowCount; row++) // Skip header
                {
                    var subscriber = new Subscriber
                    {
                        // Assuming the columns are in the order specified in the file
                        ContactFullName = worksheet.Cells[row, 3]?.Value?.ToString() ?? "",
                        AccountName = worksheet.Cells[row, 4]?.Value?.ToString() ?? "",
                        Address1 = worksheet.Cells[row, 5]?.Value?.ToString() ?? "",
                        Address2 = worksheet.Cells[row, 6]?.Value?.ToString() ?? "",
                        Address3 = worksheet.Cells[row, 7]?.Value?.ToString() ?? "",
                        City = worksheet.Cells[row, 8]?.Value?.ToString() ?? "",
                        StateProvince = worksheet.Cells[row, 9]?.Value?.ToString() ?? "",
                        PostCode = worksheet.Cells[row, 10]?.Value?.ToString() ?? "",
                        Country = worksheet.Cells[row, 11]?.Value?.ToString() ?? ""
                    };
                    subscribers.Add(subscriber);
                }
            }
        }

        private void LoadSuppliers()
        {
            suppliers = new List<Supplier>();

            // Get the project root directory
            string projectRoot = GetProjectRootDirectory();
            string filePath = Path.Combine(projectRoot, "Data", "PAGEANT ROUTING GUIDE.xlsx");

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The supplier data file was not found at: {filePath}");
            }

            // Load from PAGEANT ROUTING GUIDE.xlsx
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Skip header
                {
                    var supplier = new Supplier
                    {
                        // Assuming the first column is Country and the second is Supplier Name
                        Country = worksheet.Cells[row, 1]?.Value?.ToString() ?? "",
                        SupplierName = worksheet.Cells[row, 2]?.Value?.ToString() ?? ""
                    };
                    suppliers.Add(supplier);
                }
            }
        }

        private string GetProjectRootDirectory()
        {
            // Get the current executable directory
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;

            // Navigate up from bin/Debug/net8.0-windows to the project root
            DirectoryInfo dir = new DirectoryInfo(currentDir);

            // Go up until we find the project root (where CitipostWarehouseApplication.csproj file is located)
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

        private void AssignSuppliersToSubscribers()
        {
            foreach (var subscriber in subscribers)
            {
                // Find supplier based on country matching
                var supplier = suppliers.FirstOrDefault(s =>
                    s.Country.Equals(subscriber.Country, StringComparison.OrdinalIgnoreCase));

                if (supplier != null)
                {
                    subscriber.AssignedSupplier = supplier.SupplierName;
                }
                else
                {
                    // Default supplier for all other countries
                    subscriber.AssignedSupplier = "BP 2";
                }
            }
        }

        public List<Subscriber> GetSubscribers()
        {
            return subscribers;
        }

        public List<Supplier> GetSuppliers()
        {
            return suppliers;
        }

        public List<SummaryReport> GenerateSummaryReport()
        {
            var summary = subscribers
                .GroupBy(s => s.AssignedSupplier)
                .Select(g => new SummaryReport
                {
                    // Group by AssignedSupplier to summarize
                    SupplierName = g.Key,
                    TotalItems = g.Count(),
                    Countries = g.Select(s => s.Country).Distinct().ToList()
                })
                .ToList();

            return summary;
        }

        public void SaveProcessedData()
        {
            string projectRoot = GetProjectRootDirectory();
            string saveInDataFolder = Path.Combine(projectRoot, "Reports");

            if (!Directory.Exists(saveInDataFolder))
            {
                Directory.CreateDirectory(saveInDataFolder);
            }

            string fileName = "ProcessedWarehouseData_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            string fullPath = Path.Combine(saveInDataFolder, fileName);

            // Save updated data back to Excel with supplier assignments
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Processed Data");

                // Headers
                worksheet.Cells[1, 1].Value = "Contact Name";
                worksheet.Cells[1, 2].Value = "Account Name";
                worksheet.Cells[1, 3].Value = "Address 1";
                worksheet.Cells[1, 4].Value = "Address 2";
                worksheet.Cells[1, 5].Value = "Address 3";
                worksheet.Cells[1, 6].Value = "City";
                worksheet.Cells[1, 7].Value = "State/Province";
                worksheet.Cells[1, 8].Value = "Post Code";
                worksheet.Cells[1, 9].Value = "Country";
                worksheet.Cells[1, 10].Value = "Assigned Supplier";

                // Data
                for (int i = 0; i < subscribers.Count; i++)
                {
                    var sub = subscribers[i];
                    int row = i + 2;

                    worksheet.Cells[row, 1].Value = sub.ContactFullName;
                    worksheet.Cells[row, 2].Value = sub.AccountName;
                    worksheet.Cells[row, 3].Value = sub.Address1;
                    worksheet.Cells[row, 4].Value = sub.Address2;
                    worksheet.Cells[row, 5].Value = sub.Address3;
                    worksheet.Cells[row, 6].Value = sub.City;
                    worksheet.Cells[row, 7].Value = sub.StateProvince;
                    worksheet.Cells[row, 8].Value = sub.PostCode;
                    worksheet.Cells[row, 9].Value = sub.Country;
                    worksheet.Cells[row, 10].Value = sub.AssignedSupplier;
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                package.SaveAs(new FileInfo(fullPath));
            }
        }
    }
}
