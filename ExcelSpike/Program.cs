using ExcelSpike.Aspose;
using ExcelSpike.ClosedXML;
using ExcelSpike.Common.Abstractions;
using ExcelSpike.InteropExcel;
using ExcelSpike.OpenXML;
using ExcelSpike.Spreadsheet;
using System;
using System.Threading.Tasks;

namespace ExcelSpike
{
    class Program
    {
        public static void Main()
        {
            MainAsync().GetAwaiter().GetResult();
        }

        private static async Task MainAsync()
        {
            Console.WriteLine("Hello World!");

            //// Aspose
            //var asposeExcelService = new AsposeExcelService();
            //await RunTest(asposeExcelService, nameof(AsposeExcelService));

            //// SpreadsheetGear
            //var spreadsheetGearExcelService = new SpreadsheetGearExcelService();
            //await RunTest(spreadsheetGearExcelService, nameof(SpreadsheetGearExcelService));
            
            //// No InMemory formula calculation
            //// OpenXML
            //var openXMLExcelService = new OpenXMLExcelService();
            //await RunTest(openXMLExcelService, nameof(OpenXMLExcelService));

            //// Requite Office installation
            //// Works only with full framework
            //// Microsoft.Office.Interop.Excel
            //var interopExcelService = new InteropExcelService();
            //await RunTest(interopExcelService, nameof(InteropExcelService));

            // ClosedXML
            var closedXMLExcelService = new ClosedXMLExcelService();
            await RunTest(closedXMLExcelService, nameof(ClosedXMLExcelService));
        }

        public static async Task RunTest(IExcelService excelService, string prefix)
        {
            var dataDir = @"C:\RateSetter\spikes\excel-spike\ExcelSpike\data\";
            var tmpDataDir = $@"C:\RateSetter\spikes\excel-spike\ExcelSpike\data\{prefix}tmp\";
            var fileName = "testWorkBook.xlsx";
            var newFileName = "generatedWorkBook";

            // Test 1
            Console.WriteLine(await excelService.GetFromulaValue());

            // Test 2
            Console.WriteLine(await excelService.GetFileContent(dataDir, fileName));

            // Test 3
            await excelService.GenerateNewExcelFromTemplate(tmpDataDir, newFileName);
        } 
    }
}
