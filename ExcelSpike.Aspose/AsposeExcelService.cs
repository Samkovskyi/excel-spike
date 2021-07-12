using Aspose.Cells;
using ExcelSpike.Common.Abstractions;
using System;
using System.Threading.Tasks;

namespace ExcelSpike.Aspose
{
    public class AsposeExcelService: IExcelService
    {
        public Task GenerateNewExcelFromTemplate(string dataDir, string fileName)
        {
            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiating a Workbook object
            Workbook workbook = new Workbook();

            // Adding a new worksheet to the Excel object
            int sheetIndex = workbook.Worksheets.Add();

            // Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.Worksheets[sheetIndex];

            // Adding a value to "A1" cell
            worksheet.Cells["A1"].PutValue(1);

            // Adding a value to "A2" cell
            worksheet.Cells["A2"].PutValue(2);

            // Adding a value to "A3" cell
            worksheet.Cells["A3"].PutValue(3);

            // Adding a SUM formula to "A4" cell
            worksheet.Cells["A4"].Formula = "=SUM(A1:A3)";

            // Saving the Excel file
            workbook.Save($"{dataDir}{fileName}.xlsx", SaveFormat.Xlsx);
            // ExEnd:1

            return Task.CompletedTask;
        }

        public Task<string> GetFileContent(string dataDir, string fileName)
        {        
            // Opening Microsoft Excel 2007 Xlsx Files
            // Instantiate LoadOptions specified by the LoadFormat.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Create a Workbook object and opening the file from its path
            Workbook workbook = new Workbook($"{ dataDir }{fileName}", loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Adding a value to "A1" cell
            return Task.FromResult(worksheet.Cells["A1"].Value.ToString());
        }

        public Task<string> GetFromulaValue()
        {

            // Instantiating a Workbook object
            Workbook workbook = new Workbook();

            // Adding a new worksheet to the Excel object
            int sheetIndex = workbook.Worksheets.Add();

            // Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.Worksheets[sheetIndex];

            // Adding a value to "A1" cell
            worksheet.Cells["A1"].PutValue(1);

            // Adding a value to "A2" cell
            worksheet.Cells["A2"].PutValue(2);

            // Adding a value to "A3" cell
            worksheet.Cells["A3"].PutValue(3);

            // Adding a SUM formula to "A4" cell
            worksheet.Cells["A4"].Formula = "=SUM(A1:A3)";

            // Calculating the results of formulas
            workbook.CalculateFormula();

            // Get the calculated value of the cell
            return Task.FromResult(worksheet.Cells["A4"].Value.ToString());
        }
    }
}
