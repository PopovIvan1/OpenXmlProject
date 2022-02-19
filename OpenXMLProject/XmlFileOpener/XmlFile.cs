using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XmlFileOpener
{
    public class XmlFile
    {
        private string content = "";
        private string fileName = "";

        public XmlFile(string fileName)
        {
            this.fileName = fileName;
            setContent();
        }

        public string GetContent()
        {
            return content;
        }

        private void setContent()
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    content += getCellContent(cell, workbookPart);
                }
                content += '\n';
            }
        }

        private string getCellContent(Cell cell, WorkbookPart workbookPart)
        {
            string text;
            if (cell.CellValue != null)
            {
                text = cell.CellValue.Text;
                if (cell.DataType != null)
                {
                    SharedStringTablePart stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                    {
                        text = stringTable.SharedStringTable.ElementAt(int.Parse(text)).InnerText;
                    }
                }
                return text + ' ';
            }
            else
            {
                return "              ";
            }
        }
    }
}
