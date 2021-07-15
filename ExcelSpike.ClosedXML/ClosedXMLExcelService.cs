using ClosedXML.Excel;
using ExcelSpike.Common.Abstractions;
using System;
using System.Threading.Tasks;

namespace ExcelSpike.ClosedXML
{
    public class ClosedXMLExcelService : IExcelService
    {
        public Task GenerateNewExcelFromTemplate(string dataDir, string fileName)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs($"{dataDir}{fileName}.xlsx");
                return Task.CompletedTask;
            }
        }

        public Task<string> GetFileContent(string dataDir, string fileName)
        {
            using (var workbook = new XLWorkbook($"{dataDir}{fileName}"))
            {
                if (workbook.Worksheets.TryGetWorksheet("Sheet1", out var worksheet))
                {
                    return Task.FromResult(worksheet.Cell("A1").Value.ToString());

                }
                return Task.FromResult("failed");
            }
        }

        public Task<string> GetFromulaValue()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "1";
                worksheet.Cell("A2").Value = "2";
                worksheet.Cell("A3").Value = "3";
                worksheet.Cell("A4").FormulaA1 = "=SUM(A1:A3)";
                return Task.FromResult(worksheet.Cell("A4").Value.ToString());
            }
        }
    }
}
