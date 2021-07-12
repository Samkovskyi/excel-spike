using ExcelSpike.Common.Abstractions;
using System;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelSpike.InteropExcel
{
    public class InteropExcelService : IExcelService
    {
        public Task GenerateNewExcelFromTemplate(string dataDir, string fileName)
        {
            return Task.CompletedTask;
        }

        public Task<string> GetFileContent(string dataDir, string fileName)
        {
            //create a instance for the Excel object  
            var excelApplication = new Application();


            //pass that to workbook object  
            Workbook workBook = excelApplication.Workbooks.Open($"{dataDir}{fileName}");


            // statement get the workbookname  
            string excelWorkbookName = workBook.Name;

            // statement get the worksheet count  
            int worksheetcount = workBook.Worksheets.Count;

            Worksheet worksheet = (Worksheet)workBook.Worksheets[1];

            // statement get the firstworksheetname  

            string firstworksheetname = worksheet.Name;

            //statement get the first cell value  
            var firstcellvalue = ((Range)worksheet.Cells[1, 1]).Value;

            return Task.FromResult(firstcellvalue.ToString());
        }

        public Task<string> GetFromulaValue()
        {
            return Task.FromResult("");
        }
    }
}
