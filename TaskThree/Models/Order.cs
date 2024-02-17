using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree.files
{
    class Order
    {
        public int RequestID { get; set; }
        public int ProductID { get; set; }
        public int CustomerID { get; set; }
        public decimal Number { get; set; }
        public int ProductAmount { get; set; }
        public DateTime Date { get; set; }

    }
}
