using System;
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
            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                int index = insertSharedStringItem(content);

                WorksheetPart worksheetPart = insertWorksheet(spreadsheetDocument.WorkbookPart);

                Cell cell = insertCellInWorksheet("A", 1, worksheetPart);

                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                worksheetPart.Worksheet.Save();
            }
        }

        private int insertSharedStringItem(string content)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int index = 0;

            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == content)
                {
                    return index;
                }

                index++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(content)));
            shareStringPart.SharedStringTable.Save();

            return index;
        }

        private WorksheetPart insertWorksheet(WorkbookPart workbookPart)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        private Cell insertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            worksheet = worksheetPart.Worksheet;
            sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row = getCurrentRowFromSheetData(rowIndex);

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                return getNewCell(row, cellReference);
            }
        }

        private Row getCurrentRowFromSheetData(uint rowIndex)
        {
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                Row row = new Row() { RowIndex = rowIndex };
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
