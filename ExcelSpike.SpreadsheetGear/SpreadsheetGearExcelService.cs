using ExcelSpike.Common.Abstractions;
using System.Threading.Tasks;
using SpreadsheetGear;

namespace ExcelSpike.Spreadsheet
{
    public class SpreadsheetGearExcelService : IExcelService
    {
        public Task GenerateNewExcelFromTemplate(string dataDir, string fileName)
        {
            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiating a Workbook object
            IWorkbook workbook = Factory.GetWorkbook();
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            IRange cells = worksheet.Cells;

            //// Adding a value to "A1" cell
            cells["A1"].Value = "1";

            //// Adding a value to "A2" cell
            cells["A2"].Value = "2";
            
            //// Adding a value to "A3" cell
            cells["A3"].Value = "3";

            //// Adding a SUM formula to "A4" cell
            cells["A4"].Formula = "=SUM(A1:A3)";

            //// Saving the Excel file
            workbook.SaveAs($"{dataDir}{fileName}.xls", FileFormat.Excel8);
           
            return Task.CompletedTask;
        }

        public Task<string> GetFileContent(string dataDir, string fileName)
        {            
            //// Opening Microsoft Excel 2007 Xlsx Files
            IWorkbook workbook = Factory.GetWorkbook($"{dataDir}{fileName}");
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];        

            //// Adding a value to "A1" cell
            return Task.FromResult(worksheet.Cells["A1"].Value.ToString());
        }

        public Task<string> GetFromulaValue()
        {
            // Instantiating a Workbook object
            IWorkbook workbook = Factory.GetWorkbook();
            IWorksheet worksheet = workbook.Worksheets["Sheet1"];
            IRange cells = worksheet.Cells;

            //// Adding a value to "A1" cell
            cells["A1"].Value = "1";

            //// Adding a value to "A2" cell
            cells["A2"].Value = "2";

            //// Adding a value to "A3" cell
            cells["A3"].Value = "3";

            //// Adding a SUM formula to "A4" cell
            cells["A4"].Formula = "=SUM(A1:A3)";

            //// Get the calculated value of the cell
            return Task.FromResult(cells["A4"].Value.ToString());
        }
    }
}
