using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Reader.Models
{
    internal class Acquisition
    {
        string ID { get; set; }

        string OrderDate { get; set; }

        string Region { get; set; }

        string Rep { get; set; }

        string Item { get; set; }

        int Units { get; set; }

        float UnitCost { get; set; }

        int Total { get; set; }

    }
}
