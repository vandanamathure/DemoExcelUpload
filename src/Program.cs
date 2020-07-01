using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using CommandLine;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace VandanaMathure.DemoExcelUpload
{
    class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<CommandLineOptions>(args)
                .MapResult(ProcessFile, ParseFailed);
        }

        private static Task ParseFailed(IEnumerable<Error> errors)
        {
            foreach (var error in errors)
            {
                Console.WriteLine(error.ToString());
            }

            return Task.CompletedTask;
        }

        private static Task ProcessFile(CommandLineOptions options)
        {
            if (!File.Exists(options.FileName))
            {
                Console.WriteLine($"File {options.FileName} does not exist");
                return Task.CompletedTask;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                var workbook = excelApp.Workbooks.Open(options.FileName);
                Excel.Worksheet worksheet = workbook.Worksheets[options.WorksheetName];
                var countries = new List<Country>();
                for (var rowIndex = options.StartRowIndex; rowIndex <= options.EndRowIndex; rowIndex++)
                {
                    var country = new Country();
                    for (var colIndex = options.StartColumnIndex; colIndex <= options.EndColumnIndex; colIndex++)
                    {
                        Excel.Range range = worksheet.Cells[rowIndex, colIndex];
                        if (colIndex == options.StartColumnIndex)
                            country.Name = range.Value;
                        if (colIndex == options.EndColumnIndex)
                            country.Code = range.Value;
                    }
                    countries.Add(country);
                }

                Console.Write(JsonConvert.SerializeObject(countries, Formatting.Indented));

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(worksheet);

                // Close and Release Workbook
                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                return Task.CompletedTask;
            }
            catch (Exception)
            {
                excelApp?.Quit();
                throw;
            }
            
        }
    }
}
