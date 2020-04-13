using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RaffleBlazor.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RaffleBlazor.Utility
{
    public class ExportProvider
    {
        public async Task<string> GetXlsxAsync(IReadOnlyCollection<Department> departments)
        {
            if (departments == null)
                throw new InvalidOperationException("尚未進行選股，無法匯出。");
            if (departments.Count == 0)
                throw new InvalidOperationException("股別數為零，無法匯出。");
            if (departments.Select(x => x.Students).Sum(x => x.Count) == 0)
                throw new InvalidOperationException("股內人數加總為零，無法匯出。");

            var date = DateTime.Now;
            var ms = new MemoryStream();
            using (var spreadsheet = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Insert stylesheet to make the top row bold
                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = CreateStylesheet();

                // Define sheet
                var sheetName = $"{date:yyyy}-股員名單";
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name =  sheetName};
                sheets.AppendChild(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Insert data
                var row = new Row { RowIndex = 1 };
                row.Append(
                    new Cell { CellValue = new CellValue("學號"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("姓名"), DataType = CellValues.String },
                    new Cell { CellValue = new CellValue("股別"), DataType = CellValues.String });
                sheetData.AppendChild(row);

                uint rowIndex = 2;
                foreach (var department in departments)
                {
                    foreach (var student in department.Students)
                    {
                        var entryRow = new Row { RowIndex = rowIndex };
                        entryRow.Append(
                            new Cell { CellValue = new CellValue(student.Id), DataType = CellValues.String, StyleIndex = 1 },
                            new Cell { CellValue = new CellValue(student.Name), DataType = CellValues.String, StyleIndex = 1 },
                            new Cell { CellValue = new CellValue(student.Department.Name), DataType = CellValues.String, StyleIndex = 1 });
                        sheetData.AppendChild(entryRow);
                        rowIndex++;
                    }
                }

                // Add title
                spreadsheet.PackageProperties.Title = $"{date:yyyy} 選股結果";
                spreadsheet.PackageProperties.Creator = "RaffleBlazor 選股程式";
            }
            return await ReadStreamToBase64(ms);
        }

        // Stylesheet has to follow this order:
        //      Font -> Fills/Borders -> CellFormats
        // If you change *any* of the order, Excel will consider the spreadsheet broken.
        internal Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();
            var fonts = new Fonts();
            fonts.AppendChild(new Font
            {
                Bold = new Bold(),
                FontName = new FontName { Val = "Microsoft YaHei" },
                FontSize = new FontSize { Val = 12 },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 1 }
            });
            fonts.AppendChild(new Font
            {
                FontName = new FontName { Val = "Microsoft YaHei Light" },
                FontSize = new FontSize { Val = 12 },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 1 }
            });
            fonts.KnownFonts = true;
            fonts.Count = (uint)fonts.ChildElements.Count;
            stylesheet.AppendChild(fonts);

            // Default everything else because Excel considers this 
            // spreadsheet broken if it's missing *any* of these.
            Fill fill = new Fill() { PatternFill = new PatternFill() };
            Fills fills = new Fills();
            fills.AppendChild(fill);
            fills.Count = (uint)fills.ChildElements.Count;
            stylesheet.AppendChild(fills);

            Border border = new Border() { LeftBorder = new LeftBorder(), RightBorder = new RightBorder(), BottomBorder = new BottomBorder(), DiagonalBorder = new DiagonalBorder(), TopBorder = new TopBorder() };
            Borders borders = new Borders();
            borders.AppendChild(border);
            borders.Count = (uint)borders.ChildElements.Count;
            stylesheet.AppendChild(borders);
            
            // Now we can actually define the cell formats.
            // Screw OpenXML.
            var cellFormats = new CellFormats();
            var titleCellFormat = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 };
            var regularCellFormat = new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true };
            cellFormats.AppendChild(titleCellFormat);
            cellFormats.AppendChild(regularCellFormat);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            stylesheet.AppendChild(cellFormats);

            return stylesheet;
        }

        internal async Task<string> ReadStreamToBase64(MemoryStream stream)
        {
            Memory<byte> memory = new byte[stream.Length];
            stream.Position = 0;
            await stream.ReadAsync(memory);
            return Convert.ToBase64String(memory.Span);
        }
    }
}
