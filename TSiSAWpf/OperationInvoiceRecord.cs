using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSiSAWpf
{
    class OperationInvoiceRecord
    {
        public string FullName { get; set; }
        public string ShortName { get; set; }
        public DateTime Date { get; set; }
        public double Amount { get; set; }
        public OperationInfoType OperationInfoType { get; set; }
    }
}
