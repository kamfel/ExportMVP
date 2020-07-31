using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClosedXML
{
    class Program
    {
        static void Main(string[] args)
        {
            var dataClosedXML = GenerateDataClosedXML();

            ExcelExporter<Model>.ExportToFile("eksportClosedXML.xlsx", dataClosedXML, "closed xml");

            var dataNPOI = GenerateDataNPOI();


            ExcelExporterNPOI<ModeNPOI>.ExportToFile("eksportNPOI.xlsx", dataNPOI, "NPOI eksport");

            Console.ReadKey();
        }

        private static IEnumerable<Model> GenerateDataClosedXML()
        {
            return new List<Model>()
            {
                new Model(),
                new Model(),
                new Model(),
            };
        }

        private static IEnumerable<ModeNPOI> GenerateDataNPOI()
        {
            return new List<ModeNPOI>()
            {
                new ModeNPOI(),
                new ModeNPOI(),
                new ModeNPOI(),
            };
        }
    }
}
