using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
namespace ConverterApp
{
    internal class Program
    {
        static void Main()
        {
            var excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(@"D:\PROJECT\REPORT\test2.xlsx");
            List<string> sheetNames = new List<string>();
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                sheetNames.Add(sheet.Name);
            }
            foreach (var name in sheetNames)
            {
                Console.WriteLine(name);
            }
            //輸入您要列印的workbook 例如 ：sheet1,sheet3,...
            Console.WriteLine("workbook:");
            string input = Console.ReadLine();
            var selectedSheets = new HashSet<string>(input.Split(','));
            Excel.Workbook tempWorkbook = workbooks.Add();

            try
            {
                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    
                    sheet.Copy(Type.Missing, tempWorkbook.Sheets[tempWorkbook.Sheets.Count]);
                    if (selectedSheets.Contains(sheet.Name))
                    {
                        Excel.PageSetup pageSetup = sheet.PageSetup;
                        pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                        pageSetup.Zoom = false;
                        pageSetup.FitToPagesWide = 1;
                        pageSetup.FitToPagesTall = 1;
                        pageSetup.CenterFooter = "&P";
                        Console.WriteLine($"output file :{sheet.Name}");
                    }
                }
                Excel.Worksheet defaultSheet = (Excel.Worksheet)tempWorkbook.Sheets[1];
                defaultSheet.Delete();
                string outputPath = $@"D:\PROJECT\REPORT\outputfile\test2.pdf";
                tempWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("err: " + ex.Message);
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

      



    }
}
