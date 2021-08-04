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
        public int EmployeerCodeColumnNumber { get; set; }
        public string EdrpoSign { get; set; }
        public int EdrpoSignColumnNumber { get; set; }
        public string NameSign { get; set; }
        public int NameSignColumnNumber { get; set; }
        public string EdrpoLR { get; set; }
        public int EdrpoLRColumnNumber { get; set; }
        public string NameLR { get; set; }
        public int NameLRColumnNumber { get; set; }
        public int ActionType { get; set; }
        public int ActionTypeColumnNumber { get; set; }
        public int AttributeType { get; set; }
        public int AttributeTypeColumnNumber { get; set; }
        public DateTime ActionDT { get; set; }
        public int ActionDTColumnNumber { get; set; }
        public string ActionText { get; set; }
        public int ActionTextColumnNumber { get; set; }
        public string LeaveReason { get; set; }
        public int LeaveReasonColumnNumber { get; set; }
        public string DocType { get; set; }
        public int DocTypeColumnNumber { get; set; }
        public DateTime DocDT { get; set; }
        public int DocDTColumnNumber { get; set; }
        public string DocNumber { get; set; }
        public int DocNumberColumnNumber { get; set; }
    }
}
