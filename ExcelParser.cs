using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace LaborBookXML
{
    class ExcelParser
    {
        static List<Record> records = new List<Record>();


        public static void ParseFile(Excel.Application app)
        {
            Excel.Worksheet sheet = app.Worksheets.get_Item(1);

            for (int i = 2; i <= sheet.Rows.CurrentRegion.EntireRow.Count; i++)
            {
                Record record = new Record();
                record.EmployeerCode = sheet.Cells[i, 1].Value.ToString() ?? "";
                record.EdrpoSign = sheet.Cells[i, 2].Value.ToString() ?? "";
                record.NameSign = sheet.Cells[i, 3].Value.ToString();
                record.EdrpoLR = sheet.Cells[i, 4].Value.ToString();
                record.NameLR = sheet.Cells[i, 5].Value.ToString();
                record.ActionType = Convert.ToInt32(sheet.Cells[i, 6].Value.ToString());
                record.AttributeType = Convert.ToInt32(sheet.Cells[i, 7].Value.ToString() ?? "0");
                record.ActionDT = Convert.ToDateTime(sheet.Cells[i, 8].Value.ToString() ?? "01.01.0001");
                record.ActionText = sheet.Cells[i, 9].Value.ToString();
                record.LeaveReason = Convert.ToString((sheet.Cells[i, 10]).Value ?? "");
                record.DocType = sheet.Cells[i, 11].Value.ToString();
                record.DocDT = Convert.ToDateTime(sheet.Cells[i, 12].Value.ToString() ?? "01.01.0001");
                record.DocNumber = sheet.Cells[i, 13].Value.ToString();

                records.Add(record);

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
                final += record.ActionDT.Year + "-" + record.ActionDT.Month + "-" + record.ActionDT.Day;
                final += "</ACTION_DT><ACTION_TEXT>";
                final += record.ActionText;
                final += "</ACTION_TEXT><LEAVE_REASON>";
                final += record.LeaveReason;
                final += "</LEAVE_REASON><DOC_TYPE>";
                final += record.DocType;
                final += "</DOC_TYPE><DOC_DT>";
                final += record.DocDT.Year + "-" + record.DocDT.Month + "-" + record.DocDT.Day;
                final += "</DOC_DT><DOC_NUMBER>";
                final += record.DocNumber;
                final += "</DOC_NUMBER></RECORD>";
            }

            final += "</RECORDS></DOCUMENT>";

            return final;
        }
    }
}
