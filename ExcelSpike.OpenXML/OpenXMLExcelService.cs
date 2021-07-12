using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelSpike.Common.Abstractions;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelSpike.OpenXML
{
    public class OpenXMLExcelService : IExcelService
    {
        public Task GenerateNewExcelFromTemplate(string dataDir, string fileName)
        {
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            var filepath = $"{dataDir}{fileName}.xlsx";
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var cellA1 = InsertCellInWorksheet("A", 1, worksheetPart);
            cellA1.CellValue = new CellValue("1");
            cellA1.DataType = new EnumValue<CellValues>(CellValues.Number);
            var cellA2 = InsertCellInWorksheet("A", 2, worksheetPart);
            cellA2.CellValue = new CellValue("2");
            cellA2.DataType = new EnumValue<CellValues>(CellValues.Number);
            var cellA3 = InsertCellInWorksheet("A", 3, worksheetPart);
            cellA3.CellValue = new CellValue("3");
            cellA3.DataType = new EnumValue<CellValues>(CellValues.Number);
            var cellA4 = InsertCellInWorksheet("A", 4, worksheetPart);
            cellA4.CellFormula = new CellFormula("=SUM(A1:A3)");

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

            return Task.CompletedTask;
        }

        public Task<string> GetFileContent(string dataDir, string fileName)
        {
            return Task.FromResult(GetCellValue($"{dataDir}{fileName}", "Sheet1", "A1"));
        }

        public Task<string> GetFromulaValue()
        {
            using (var memoryStream = new MemoryStream())
            {
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook(); ;

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                var cellA1 = InsertCellInWorksheet("A", 1, worksheetPart);
                cellA1.CellValue = new CellValue("1");
                cellA1.DataType = new EnumValue<CellValues>(CellValues.Number);
                var cellA2 = InsertCellInWorksheet("A", 2, worksheetPart);
                cellA2.CellValue = new CellValue("2");
                cellA2.DataType = new EnumValue<CellValues>(CellValues.Number);
                var cellA3 = InsertCellInWorksheet("A", 3, worksheetPart);
                cellA3.CellValue = new CellValue("3");
                cellA3.DataType = new EnumValue<CellValues>(CellValues.Number);
                var cellA4 = InsertCellInWorksheet("A", 4, worksheetPart); 
                cellA4.DataType = new EnumValue<CellValues>(CellValues.Number);
                cellA4.CellFormula = new CellFormula("=SUM(A1:A3)");
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties = new CalculationProperties { ForceFullCalculation = true, CalculationOnSave = true, FullCalculationOnLoad = true };
                spreadsheetDocument.WorkbookPart.Workbook.Save();
                spreadsheetDocument.Save();
                spreadsheetDocument.WorkbookPart.Workbook.Reload();
                var cellValue = spreadsheetDocument.WorkbookPart.WorksheetParts
                    .SelectMany(part => part.Worksheet.Elements<SheetData>())
                    .SelectMany(data => data.Elements<Row>())
                    .SelectMany(row => row.Elements<Cell>())
                    .Where(cell => cell.CellReference == "A4")
                    .Where(cell => cell.CellValue != null)
                    .FirstOrDefault(); 
                
                return Task.FromResult(cellA4.CellValue?.ToString());
            }
        }

        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell.InnerText.Length > 0)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}
