using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelRefreshAll
{
    class Program
    {
        static void Main(string[] args)
        {
            
            try
            {
                Console.WriteLine("Refreshing knowledgecenter connection by opening explorer window...");

                Process.Start(@"\\knowledgecenter.eqt.com@SSL\DavWWWRoot\it\home\gis\GIS_Portal\Shared Documents\Application_Documents");
                string inputWorkbookPath = @"\\knowledgecenter.eqt.com@SSL\DavWWWRoot\it\home\gis\GIS_Portal\Shared Documents\Application_Documents\Stats_Test.xlsx";

                // The Microsoft.Office.Interop.Excel.Application interface represents the entire Microsoft Excel application
                Application excelApplication = new Microsoft.Office.Interop.Excel.Application();

                // Suppress alert windows
                excelApplication.DisplayAlerts = false;

                // Opens a workbook
                Workbook workbook = excelApplication.Workbooks.Open(inputWorkbookPath);

                // Refreshes all external data ranges and PivotTable reports in the specified workbook
                workbook.RefreshAll();

                // Suspends the current thread for the specified number of milliseconds (1000 ms = 1 s)
                Console.WriteLine("Pausing to allow data to refresh...");
                System.Threading.Thread.Sleep(20000);

                // Save the workbook
                workbook.Save();

                // Close the workbook
                workbook.Close(false, inputWorkbookPath, null);

                // Quit the application
                excelApplication.Quit();

                // Decrement the reference count of the specified Runtime Callable Wrapper associated with the specified COM object
                // Component Object Model (COM) is a platform-independent, distributed, object-oriented system for creating binary 
                // software components that can interact.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

                workbook = null;

            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred.");
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press the enter key to exit.");
                Console.Read();
            }

        }
    }
}
