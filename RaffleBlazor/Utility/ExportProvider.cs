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
                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = CreateStylesheet();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = $"{date:yyyy}-股員名單" };
                sheets.AppendChild(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

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
                            new Cell { CellValue = new CellValue(student.Id), DataType = CellValues.String },
                            new Cell { CellValue = new CellValue(student.Name), DataType = CellValues.String },
                            new Cell { CellValue = new CellValue(student.Department.Name), DataType = CellValues.String });
                        sheetData.AppendChild(entryRow);
                        rowIndex++;
                    }
                }
            }
            return await ReadStreamToBase64(ms);
        }

        internal Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();
            var fonts = new Fonts();
            fonts.AppendChild(new Font
            {
                FontName = new FontName { Val = "Microsoft YaHei" },
                FontSize = new FontSize { Val = 12 },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
            });
            fonts.Count = (uint)fonts.ChildElements.Count;
            stylesheet.AppendChild(fonts);
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
