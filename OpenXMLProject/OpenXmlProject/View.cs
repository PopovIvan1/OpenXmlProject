using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms;

namespace OpenXmlProject
{
    public partial class View : Form
    {
        public View()
        {
            InitializeComponent();
            openToolStripMenuItem.Click += selectXmlFile;
        }

        private void selectXmlFile(object theSender, EventArgs theArgs)
        {
            myRichTextBox.Clear();
            OpenFileDialog aFileDialog = new OpenFileDialog();
            aFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (aFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (SpreadsheetDocument aSpreadsheetDocument = SpreadsheetDocument.Open(aFileDialog.FileName, false))
                {
                    WorkbookPart aWorkbookPart = aSpreadsheetDocument.WorkbookPart;
                    WorksheetPart aWorksheetPart = aWorkbookPart.WorksheetParts.First();
                    SheetData sheetData = aWorksheetPart.Worksheet.Elements<SheetData>().First();
                    string aText;
                    foreach (Row aRow in sheetData.Elements<Row>())
                    {
                        foreach (Cell aCell in aRow.Elements<Cell>())
                        {
                            if (aCell.CellValue != null)
                            {
                                aText = aCell.CellValue.Text;
                                myRichTextBox.Text += aText + ' ';
                            }
                        }
                        myRichTextBox.Text += '\n';
                    }
                }
            }
        }
    }
}
