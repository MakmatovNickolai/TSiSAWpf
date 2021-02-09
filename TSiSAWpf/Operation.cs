using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSiSAWpf
{
    class Operation
    {
        public int Number { get; set; }
        public DateTime Date { get; set; }
        public string[] Documents { get; set; }
        public string[] Subjects { get; set; }
        public string[] Debets { get; set; }
        public string[] Credits { get; set; }
        public double[] Sums { get; set; }

        override public string ToString()
        {
            return $"Number - {Number}, Date - {Date.ToShortDateString()} " +
                $"Documents - {string.Join(", ", Documents)} " +
                $"Subjects - {string.Join(", ", Subjects)} " +
                $"Debets - {string.Join(", ", Debets)} " +
                $"Credits - {string.Join(", ", Credits)} " +
                $"Sums - {string.Join(", ", Sums)}";
        }

    }
}
