using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LaborBookXML
{
    class ExcelParser
    {
        public static List<Record> records = new List<Record>();

        static Excel.Worksheet sheet;

         public static Dictionary<string, int> namePositionPairs = new Dictionary<string, int>();
        public static void ParseFile(Excel.Application app)
        {
            sheet = app.Worksheets.get_Item(1);
            records.Clear();
            try
            {
                for (int i = 2; i <= sheet.Rows.CurrentRegion.EntireRow.Count; i++)
                {
                    Record record = new Record();

                    record.EmployeerCode = sheet.Cells[i, namePositionPairs["EmployeerCode"]].Value.ToString() ?? "";
                    record.EdrpoSign = sheet.Cells[i, namePositionPairs["EdrpoSign"]].Value.ToString();
                    record.NameSign = sheet.Cells[i, namePositionPairs["NameSign"]].Value2.ToString();
                    record.EdrpoLR = sheet.Cells[i, namePositionPairs["EdrpoLR"]].Value.ToString();
                    record.NameLR = sheet.Cells[i, namePositionPairs["NameLR"]].Value.ToString();
                    record.ActionType = Convert.ToInt32(sheet.Cells[i, namePositionPairs["ActionType"]].Value.ToString());
                    record.AttributeType = Convert.ToInt32(sheet.Cells[i, namePositionPairs["AttributeType"]].Value.ToString());
                    record.ActionDT = Convert.ToDateTime(sheet.Cells[i, namePositionPairs["ActionDT"]].Value.ToString());
                    record.ActionText = sheet.Cells[i, namePositionPairs["ActionText"]].Value.ToString();
                    record.LeaveReason = Convert.ToString(sheet.Cells[i, namePositionPairs["LeaveReason"]].Value2 ?? "");
                    record.DocType = sheet.Cells[i, namePositionPairs["DocType"]].Value.ToString();
                    record.DocDT = Convert.ToDateTime(sheet.Cells[i, namePositionPairs["DocDT"]].Value);
                    record.DocNumber = sheet.Cells[i, namePositionPairs["DocNumber"]].Value.ToString();

                    records.Add(record);
                }
            }
            catch
            {
                MessageBox.Show("Неверные данные, проверте правильность ввода данных");
                
            }
        }

        public static string GetXmlDocuent()
        {
            string final = "";

            final += "<DOCUMENT><DOCUMENT_HEAD><KSS/><VERSION>1</VERSION><DOCUMENT_DT>" + System.DateTime.Now + "</DOCUMENT_DT></DOCUMENT_HEAD><RECORDS>";

            foreach (var record in records)
            {
                final += "<RECORD><EMPLOYER_CODE>";
                final += record.EmployeerCode;
                final += "</EMPLOYER_CODE><EDRPO_SIGN>";
                final += record.EdrpoSign;
                final += "</EDRPO_SIGN><NAME_SIGN>";
                final += record.NameSign;
                final += "</NAME_SIGN><EDRPO_LR>";
                final += record.EdrpoLR;
                final += "</EDRPO_LR><NAME_LR>";
                final += record.NameLR;
                final += "</NAME_LR><ACTION_TYPE>";
                final += record.ActionType;
                final += "</ACTION_TYPE><ATTRIBUTE_TYPE>";
                final += record.AttributeType;
                final += "</ATTRIBUTE_TYPE><ACTION_DT>";
                final += record.ActionDT.Year + "-" +
                        (record.ActionDT.Month < 10 ? "0" + record.ActionDT.Month : record.ActionDT.Month.ToString()) + "-" +
                        (record.ActionDT.Day < 10 ? "0" + record.ActionDT.Day : record.ActionDT.Day.ToString());
                final += "</ACTION_DT><ACTION_TEXT>";
                final += record.ActionText;
                final += "</ACTION_TEXT><LEAVE_REASON>";
                final += record.LeaveReason;
                final += "</LEAVE_REASON><DOC_TYPE>";
                final += record.DocType;
                final += "</DOC_TYPE><DOC_DT>";
                final += record.DocDT.Year + "-" +
                        (record.DocDT.Month < 10 ? "0" + record.DocDT.Month : record.DocDT.Month.ToString()) + "-" +
                        (record.DocDT.Day < 10 ? "0" + record.DocDT.Day : record.DocDT.Day.ToString());
                final += "</DOC_DT><DOC_NUMBER>";
                final += record.DocNumber;
                final += "</DOC_NUMBER></RECORD>";
            }

            final += "</RECORDS></DOCUMENT>";

            return final;
        }
        public static bool FindHeaders(Excel.Application app)
        {
            sheet = app.Worksheets.get_Item(1);
            //MessageBox.Show(sheet.Range["A1:M1"].Find("Employeer_Code").Column.ToString());
            //Record.EmployeerCodeColumn = Convert.ToChar(sheet.Range["A1:M1"].Find("Employeer_Code").Address.Substring(1, 1));

            if (sheet.Range["A1:M1"].Find("Employeer_Code") != null &&
               sheet.Range["A1:M1"].Find("Edrpo_Sign") != null &&
               sheet.Range["A1:M1"].Find("Name_Sign") != null &&
               sheet.Range["A1:M1"].Find("Edrpo_LR") != null &&
               sheet.Range["A1:M1"].Find("Name_LR") != null &&
               sheet.Range["A1:M1"].Find("Action_Type") != null &&
               sheet.Range["A1:M1"].Find("Attribute_Type") != null &&
               sheet.Range["A1:M1"].Find("Action_DT") != null &&
               sheet.Range["A1:M1"].Find("Action_Text") != null &&
               sheet.Range["A1:M1"].Find("Leave_Reason") != null &&
               sheet.Range["A1:M1"].Find("Doc_Type") != null &&
               sheet.Range["A1:M1"].Find("Doc_DT") != null &&
               sheet.Range["A1:M1"].Find("Doc_Number") != null
               )
            {
                namePositionPairs.Add("EmployeerCode", sheet.Range["A1:M1"].Find("Employeer_Code").Column);
                namePositionPairs.Add("EdrpoSign", sheet.Range["A1:M1"].Find("Edrpo_Sign").Column);
                namePositionPairs.Add("NameSign", sheet.Range["A1:M1"].Find("Name_Sign").Column);
                namePositionPairs.Add("EdrpoLR", sheet.Range["A1:M1"].Find("Edrpo_LR").Column);
                namePositionPairs.Add("NameLR", sheet.Range["A1:M1"].Find("Name_LR").Column);
                namePositionPairs.Add("ActionType", sheet.Range["A1:M1"].Find("Action_Type").Column);
                namePositionPairs.Add("AttributeType", sheet.Range["A1:M1"].Find("Attribute_Type").Column);
                namePositionPairs.Add("ActionDT", sheet.Range["A1:M1"].Find("Action_DT").Column);
                namePositionPairs.Add("ActionText", sheet.Range["A1:M1"].Find("Action_Text").Column);
                namePositionPairs.Add("LeaveReason", sheet.Range["A1:M1"].Find("Leave_Reason").Column);
                namePositionPairs.Add("DocType", sheet.Range["A1:M1"].Find("Doc_Type").Column);
                namePositionPairs.Add("DocDT", sheet.Range["A1:M1"].Find("Doc_DT").Column);
                namePositionPairs.Add("DocNumber", sheet.Range["A1:M1"].Find("Doc_Number").Column);

                return true;
            }
            else
            { 
                MessageBox.Show("Один или несколько заголовков столбцов не были найдены. Проверте их правильность написания и размещения. Они должны быть распаложены в диапазоне А1:М1");
                return false;
            }
        }
    }

}
