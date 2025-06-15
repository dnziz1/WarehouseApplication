using OfficeOpenXml;

namespace CitipostWarehouseApplication
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Enable visual styles for the application
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                // Run the main form of the application
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                // Log the exception (consider using a logging framework)
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}