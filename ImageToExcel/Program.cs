using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ImageToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var filename = "Image.xlsx";

            var spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Image"
            };
            sheets.Append(sheet);

            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            for (uint row = 1; row <= 10; row++)
            {
                var r = new Row() { RowIndex = row };
                sheetData.Append(r);

                Cell previousCell = null;
                for (uint col = 1; col <= 10; col++)
                {
                    char columnName = Convert.ToChar(col + 64);

                    Cell newCell = new Cell() { CellReference = $"{columnName}{row}" };
                    r.InsertAfter(newCell, previousCell);
                    previousCell = newCell;

                    newCell.CellValue = new CellValue(newCell.CellReference);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }

            worksheet.Save();
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();

            Console.WriteLine("Done.");
        }
    }
}
