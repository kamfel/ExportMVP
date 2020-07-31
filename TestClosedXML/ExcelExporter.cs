using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TestClosedXML
{


    public static class ExcelExporter<T> 
        where T : class
    {
        public static bool ExportToFile(string filepath, IEnumerable<T> data, string name = "test name")
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    CreateContent(workbook, data, name);
                    workbook.SaveAs(filepath);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

        private static void CreateContent(XLWorkbook workbook, IEnumerable<T> data, string name)
        {
            var worksheet = workbook.Worksheets.Add(name);
            var dataCount = data.Count();
            var properties = typeof(T).GetProperties();
            var propertiesCount = properties.Count();

            // To jest ciekawe :D
            // W jednej linijce cały excel wypełniony i trzeba dodać formatowanie tylko
            worksheet.Cell(1, 1).InsertTable(data, true);

            // Formatowanie
            for (int i = 1; i <= propertiesCount; i++)
            {
                var column = worksheet.Column(i);
                SetFormatting(column, properties[i - 1], dataCount);
            }
        }

        private static void SetFormatting(IXLColumn column, PropertyInfo property, int dataCount)
        {
            var headerCell = column.Cell(1).AsRange();
            var range = column.Cells(2, dataCount + 1);

            // Do tego trzeba utworzyć atrybut
            if (property.IsType<DateTime>())
            {
                range.Style.NumberFormat.NumberFormatId = (int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes; // Ten format zależy od środowiska i w PL będzie dd.MM.yyyy
            }
            else if (property.IsType<decimal>())
            {
                range.Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);
            }
        }

        private static void CreateHeaderRow(IXLRow row, IEnumerable<PropertyInfo> properties)
        {
            var names = properties.Select(x => x.Name);

            row.SetValue(names);

            row.SetAutoFilter(true);
        }

        
    }
}
