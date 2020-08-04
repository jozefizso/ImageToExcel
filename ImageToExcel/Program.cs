using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

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
            var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            worksheetPart.Worksheet = new Worksheet(
                new SheetViews(),
                new Columns(),
                new SheetData()
            );
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Image"
            };
            sheets.Append(sheet);

            var worksheet = worksheetPart.Worksheet;

            var sheetViews = worksheet.GetFirstChild<SheetViews>();
            var sheetView = new SheetView()
            {
                ShowGridLines = false,
                ShowRowColHeaders = false,
                ShowRuler = false,
                TabSelected = true,
                ZoomScaleNormal = 100,
                WorkbookViewId = 0
            };
            sheetViews?.Append(sheetView);

            var styleSheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            styleSheet.Fills = new Fills() { Count = 0 };
            styleSheet.CellFormats = new CellFormats() { Count = 0 };

            var columns = worksheet.GetFirstChild<Columns>();
            columns?.Append(new Column() { Min = 1, Max = 255, Width = 2, CustomWidth = true });

            var sheetData = worksheet.GetFirstChild<SheetData>();

            for (uint row = 1; row <= 20; row++)
            {
                var r = new Row() { RowIndex = row, Height = 10, CustomHeight = true };
                sheetData.Append(r);

                Cell previousCell = null;
                for (uint col = 1; col <= 20; col++)
                {
                    var rc = row * 10;
                    var gc = col * 10;
                    var bc = row + col * 10;
                    if (bc > 255)
                    {
                        bc = 256;
                    }

                    var color = $"FF{rc:X2}{gc:X2}{bc:X2}";
                    //var color = "FFFFFF00";
                    var bgcolor = new HexBinaryValue(color);

                    Fill fill = new Fill()
                    {
                        PatternFill = new PatternFill
                        {
                            ForegroundColor = new ForegroundColor { Rgb = bgcolor },
                            //BackgroundColor = new BackgroundColor { Rgb = "FFCC8844" },
                            PatternType = PatternValues.Solid
                        }
                    };

                    styleSheet.Fills.Append(fill);
                    var fid = styleSheet.Fills.Count++;

                    var cellFormat = new CellFormat();
                    cellFormat.FillId = fid;
                    styleSheet.CellFormats.Append(cellFormat);

                    var styleId = styleSheet.CellFormats.Count++;


                    char columnName = Convert.ToChar(col + 64);

                    Cell newCell = new Cell() 
                    {
                        CellReference = $"{columnName}{row}",
                        StyleIndex = styleId
                    };
                    r.InsertAfter(newCell, previousCell);
                    previousCell = newCell;

                    //newCell.CellValue = new CellValue(newCell.CellReference);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }

            styleSheet.Save();
            worksheet.Save();
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();

            Console.WriteLine("Done.");
        }
    }
}
