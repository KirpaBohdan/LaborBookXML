using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


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
                listBox1.Items.Clear();

                if (ExcelParser.FindHeaders(xlApp))
                {
                    ExcelParser.ParseFile(xlApp);
                    xText = ExcelParser.GetXmlDocuent();

                    if (ExcelParser.records.Count != 0)
                    {
                        button2.Enabled = true;
                        richTextBox1.Text = xText;
                        listBox1.Items.Add("Весь файл");
                        foreach (var record in ExcelParser.records)
                            listBox1.Items.Add(record.ActionText);
                    }
                }
                else
                {
                    button2.Enabled = false;
                    richTextBox1.Text = "";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            //File.WriteAllText("MyXmlFile", xText);

            saveFileDialog.Filter = "xml file(*.xml)|*.xml";
            saveFileDialog.OverwritePrompt = true;

            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                File.WriteAllText(saveFileDialog.FileName, xText);
                MessageBox.Show("Файл сохранен!");
            }
        }

        private void listBox1_OnItemClick(object sender, EventArgs e)
        {

            foreach (var record in ExcelParser.records)
            {
                if (record.ActionText == (sender as ListBox).SelectedItem.ToString())
                {
                    string textRecord = "";
                    textRecord += "<RECORD>\n<EMPLOYER_CODE>";
                    textRecord += record.EmployeerCode;
                    textRecord += "</EMPLOYER_CODE>\n<EDRPO_SIGN>";
                    textRecord += record.EdrpoSign;
                    textRecord += "</EDRPO_SIGN>\n<NAME_SIGN>";
                    textRecord += record.NameSign;
                    textRecord += "</NAME_SIGN>\n<EDRPO_LR>";
                    textRecord += record.EdrpoLR;
                    textRecord += "</EDRPO_LR>\n<NAME_LR>";
                    textRecord += record.NameLR;
                    textRecord += "</NAME_LR>\n<ACTION_TYPE>";
                    textRecord += record.ActionType;
                    textRecord += "</ACTION_TYPE>\n<ATTRIBUTE_TYPE>";
                    textRecord += record.AttributeType;
                    textRecord += "</ATTRIBUTE_TYPE>\n<ACTION_DT>";
                    textRecord += record.ActionDT.Year + "-" + 
                        (record.ActionDT.Month < 10 ? "0" + record.ActionDT.Month : record.ActionDT.Month.ToString()) + "-" + 
                        (record.ActionDT.Day < 10? "0" + record.ActionDT.Day : record.ActionDT.Day.ToString());
                    textRecord += "</ACTION_DT>\n<ACTION_TEXT>";
                    textRecord += record.ActionText;
                    textRecord += "</ACTION_TEXT>\n<LEAVE_REASON>";
                    textRecord += record.LeaveReason;
                    textRecord += "</LEAVE_REASON>\n<DOC_TYPE>";
                    textRecord += record.DocType;
                    textRecord += "</DOC_TYPE>\n<DOC_DT>";
                    textRecord += record.DocDT.Year + "-" +
                        (record.DocDT.Month < 10 ? "0" + record.DocDT.Month : record.DocDT.Month.ToString()) + "-" +
                        (record.DocDT.Day < 10 ? "0" + record.DocDT.Day : record.DocDT.Day.ToString());
                    textRecord += "</DOC_DT>\n<DOC_NUMBER>";
                    textRecord += record.DocNumber;
                    textRecord += "</DOC_NUMBER>\n</RECORD>";

                    richTextBox1.Text = textRecord;
                }
                else if ((sender as ListBox).SelectedItem.ToString() == "Весь файл")
                    richTextBox1.Text = xText;
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
