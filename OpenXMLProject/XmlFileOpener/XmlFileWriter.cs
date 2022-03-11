using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XmlFileOpener
{
    public class XmlFileWriter
    {
        private string fileName;
        private SharedStringTablePart shareStringPart;
        private SheetData sheetData;
        private Worksheet worksheet;

        public XmlFileWriter(string fileName)
        {
            this.fileName = fileName;
        }

        public void SetContent(string content)
        {
            using (var spreadsheetDocument = SpreadsheetDocument
                .Open(fileName, true))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var stringTablePart = workbookPart
                        .GetPartsOfType<SharedStringTablePart>();
                if (stringTablePart.Count() > 0)
                {
                    shareStringPart = stringTablePart?.First();
                }
                else
                {
                    shareStringPart = workbookPart
                        .AddNewPart<SharedStringTablePart>();
                }

                int index = insertSharedStringItem(content);
                WorksheetPart worksheetPart = insertWorksheet(workbookPart);

                Cell cell = insertCellInWorksheet("A", 1, worksheetPart);
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                worksheetPart.Worksheet.Save();
            }
        }

        private int insertSharedStringItem(string content)
        {
            var sharedStringTable = shareStringPart.SharedStringTable;
            if (sharedStringTable == null)
            {
                sharedStringTable = new SharedStringTable();
            }

            int index = 0;

            foreach (SharedStringItem item in sharedStringTable
                .Elements<SharedStringItem>())
            {
                if (item.InnerText == content)
                {
                    return index;
                }

                index++;
            }

            var sharedSrtingItem = new SharedStringItem(new Text(content));
            sharedStringTable.AppendChild(sharedSrtingItem);
            sharedStringTable.Save();

            return index;
        }

        private WorksheetPart insertWorksheet(WorkbookPart workbookPart)
        {
            WorksheetPart newWorksheetPart = workbookPart
                .AddNewPart<WorksheetPart>();

            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart
                .GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>()
                    .Select(sheet => sheet.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;
            Sheet newSheet = new Sheet() 
            { 
                Id = relationshipId, SheetId = sheetId, Name = sheetName 
            };
            sheets.Append(newSheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        private Cell insertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            worksheet = worksheetPart.Worksheet;
            sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row = getCurrentRowFromSheetData(rowIndex);
            var cellsWithCurrentReference = row.Elements<Cell>()
                .Where(cell => cell.CellReference.Value == cellReference);

            if (cellsWithCurrentReference.Count() > 0)
            {
                return cellsWithCurrentReference.First();
            }
            else
            {
                return getNewCell(row, cellReference);
            }
        }

        private Row getCurrentRowFromSheetData(uint rowIndex)
        {
            var rows = sheetData.Elements<Row>();
            var rowsWithCurrentIndex = rows
                .Where(row => row.RowIndex == rowIndex);

            if (rowsWithCurrentIndex.Count() > 0)
            {
                return rowsWithCurrentIndex.First();
            }
            else
            {
                Row row = new Row() 
                { 
                    RowIndex = rowIndex 
                };
                sheetData.Append(row);
                return row;
            }
        }

        private Cell getNewCell(Row row, string cellReference)
        {
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(cell
                        .CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            Cell newCell = new Cell() 
            { 
                CellReference = cellReference 
            };

            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }
}
