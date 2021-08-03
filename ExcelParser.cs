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
        public static void ParseFile(Excel.Application app)
        {
            Excel.Worksheet sheet = app.Worksheets.get_Item(1);
        }
    }
}
