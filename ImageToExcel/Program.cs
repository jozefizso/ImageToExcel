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

            var fonts = new Fonts(new Font(new FontSize() { Val = 10 }));
            var fills = new Fills(new Fill(new PatternFill() { PatternType = PatternValues.None })) { Count = 1 };
            var borders = new Borders(new Border()) { Count = 1 };
            var cellFormats = new CellFormats(new CellFormat()) { Count = 1 };
            var stylesheet = new Stylesheet(fonts, fills, borders, cellFormats);
            stylesPart.Stylesheet = stylesheet;

            var columns = worksheet.GetFirstChild<Columns>();
            columns?.Append(new Column() { Min = 1, Max = 255, Width = 1, CustomWidth = true });

            var sheetData = worksheet.GetFirstChild<SheetData>();

            for (uint row = 1; row <= 100; row++)
            {
                var r = new Row() { RowIndex = row, Height = 4, CustomHeight = true };
                sheetData.Append(r);

                Cell previousCell = null;
                for (uint col = 1; col <= 100; col++)
                {
                    var rc = col * 2;
                    var gc = row * 2;
                    var bc = row + col;
                    if (gc > 255)
                    {
                        gc = 256;
                    }
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

                    stylesheet.Fills.Append(fill);
                    var fid = stylesheet.Fills.Count++;

                    var cellFormat = new CellFormat();
                    cellFormat.FillId = fid;
                    cellFormat.ApplyFill = true;
                    stylesheet.CellFormats.Append(cellFormat);

                    var styleId = stylesheet.CellFormats.Count++;


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

            stylesheet.Save();
            worksheet.Save();
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();

            Console.WriteLine("Done.");
        }
    }
}
