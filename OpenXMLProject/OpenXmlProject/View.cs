using System;
using System.Windows.Forms;
using XmlFileOpener;

namespace OpenXmlProject
{
    public partial class View : Form
    {
        public View()
        {
            InitializeComponent();
            openToolStripMenuItem.Click += selectXmlFile;
        }

        private void selectXmlFile(object sender, EventArgs args)
        {
            richTextBox.Clear();

            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = fileDialog.FileName;

                XmlFileWriter xmlFileWriter = new XmlFileWriter(fileName);
                xmlFileWriter.SetContent("Test content");

                XmlFileReader xmlFileReader = new XmlFileReader(fileName);
                richTextBox.Text = xmlFileReader.GetContent();
            }
        }
    }
}
