﻿using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XmlFileOpener
{
    public class XmlFile
    {
        private string content = "";

        private WorkbookPart workbookPart;
        private SheetData sheetData;

        public XmlFile(string fileName)
        {
            initialize(fileName);
        }

        private void initialize(string fileName)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                setContent();
            }
        }

        private void setContent()
        {
            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    content += getCellContent(cell);
                }
                content += '\n';
            }
        }

        private string getCellContent(Cell cell)
        {
            if (cell.CellValue != null)
            {
                string text = cell.CellValue.Text;
                if (cell.DataType != null)
                {
                    text = getContentFromCellWithNotGeneralDataType(cell);
                }
                return text + ' ';
            }
            else
            {
                return "              ";
            }
        }

        private string getContentFromCellWithNotGeneralDataType(Cell cell)
        {
            SharedStringTablePart stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (stringTable != null)
            {
                return stringTable.SharedStringTable.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public string GetContent()
        {
            return content;
        }
    }
}
