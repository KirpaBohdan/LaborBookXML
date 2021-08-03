using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LaborBookXML
{
    class ExcelParser
    {
        static List<Record> records = new List<Record>();
        public static void ParseFile(Excel.Application app)
        {
            Excel.Worksheet sheet = app.Worksheets.get_Item(1);

            //List<Record> records = new List<Record>();

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
                record.LeaveReason = Convert.ToString((sheet.Cells[i, 10]).Value2 ?? "");//sheet.Cells[i, 10].Value.ToString() ?? "";
                record.DocType = sheet.Cells[i, 11].Value.ToString();
                record.DocDT = Convert.ToDateTime(sheet.Cells[i, 12].Value.ToString() ?? "01.01.0001");
                record.DocNumber = sheet.Cells[i, 13].Value.ToString();

                MessageBox.Show(record.DocDT.ToString());

                records.Add(record);
                //MessageBox.Show(sheet.Cells[1, 3].Value.ToString());
            }
        }
    }
}
