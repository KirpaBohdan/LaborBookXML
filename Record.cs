using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborBookXML
{
    class Record
    {
        public string EmployeerCode { get; set; }
        public string EdrpoSign { get; set; }
        public string NameSign { get; set; }
        public string EdrpoLR { get; set; }
        public string NameLR { get; set; }
        public int ActionType { get; set; }
        public int AttributeType { get; set; }
        public DateTime ActionDT { get; set; }
        public string ActionText { get; set; }
        public string LeaveReason { get; set; }
        public string DocType { get; set; }
        public DateTime DocDT { get; set; }
        public string DocNumber { get; set; }
    }
}
