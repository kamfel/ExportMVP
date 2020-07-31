using ClosedXML.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClosedXML
{
    // Trzeba podawać kolejność kolumn bo inaczej jest losowo.
    public class Model
    {
        [XLColumn(Header = "Header int", Order = 1)]
        public int Testint { get; set; } = 3;
        [XLColumn(Header = "Header double", Order = 2)]
        public double TestDouble { get; set; } = 2.345d;
        [XLColumn(Header = "Header float", Order = 3)]
        public float TestFlaot { get; set; } = 342.5f;
        [XLColumn(Header = "Header string", Order = 4)]
        public string TestString { get; set; } = "halo halo";
        [XLColumn(Header = "Date test", Order = 5)]
        public DateTime TestDate { get; set; } = DateTime.Now;
        [XLColumn(Header = "Header decimal", Order = 6)]
        public decimal TestDecimal { get; set; } = 25.54325m;

        [XLColumn(Ignore = true)]
        public decimal TestDecimal1 { get; set; } = 4151.324m;

    }
}
