using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSiSAWpf
{
    public class Invoice
    {
        public string ShortName { get; set; }
        public string Description { get; set; }
        public double DebetSum { get;set; }
        public double CreditSum { get; set; }
        public InvoiceType Type { get; set; }
        public double SaldoEndDebet { get; set; }
        public double SaldoEndCredit { get; set; }

    }
}
