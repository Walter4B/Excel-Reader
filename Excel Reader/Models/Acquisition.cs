using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Reader.Models
{
    internal class Acquisition
    {
        public string OrderDate { get; set; }

        public string Region { get; set; }

        public string Rep { get; set; }

        public string Item { get; set; }

        public int Units { get; set; }

        public float UnitCost { get; set; }

        public int Total { get; set; }

    }
}
