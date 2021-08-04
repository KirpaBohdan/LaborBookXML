using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace LaborBookXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string xText;

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.DefaultExt = ".xlsx";
            openFileDialog.FileName = "";
            openFileDialog.AddExtension = true;
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;";
            openFileDialog.ShowDialog();

            if (openFileDialog.FileName != "")
            {
                labelFileDirectory.Text = openFileDialog.FileName;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(openFileDialog.FileName);

                ExcelParser.namePositionPairs.Clear();

                if (ExcelParser.FindHeaders(xlApp))
                {
                    ExcelParser.ParseFile(xlApp);
                    xText = ExcelParser.GetXmlDocuent();

                    if (ExcelParser.records.Count != 0)
                        button2.Enabled = true;
                }
                else
                    button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            File.WriteAllText("MyXmlFile", xText);

            saveFileDialog.Filter = "xml file(*.xml)|*.xml";
            saveFileDialog.OverwritePrompt = true;

            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                File.WriteAllText(saveFileDialog.FileName, xText);
                MessageBox.Show("Файл сохранен!");
            }
        }
    }
}
