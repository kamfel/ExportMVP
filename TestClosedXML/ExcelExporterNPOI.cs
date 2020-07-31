using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Reflection;
using NPOI.HSSF.UserModel;

namespace TestClosedXML
{
    public static class ExcelExporterNPOI<T>
    {

        public static bool ExportToFile(string filepath, IEnumerable<T> data, string sheetName = "test name")
        {
            try
            {
                var workbook = new XSSFWorkbook();
                CreateContent(workbook, data, sheetName);

                using (var file = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

        private static void CreateContent(XSSFWorkbook workbook, IEnumerable<T> data, string name)
        {
            XSSFSheet sheet = workbook.CreateSheet(name) as XSSFSheet;
            var properties = (typeof(T)).GetProperties();
            var dataCount = data.Count();
            var propertiesCount = properties.Count();

            CreateHeaderRow(sheet.CreateRow(0), properties);

            for (int i = 1; i <= dataCount; i++)
            {
                var row = sheet.CreateRow(i);
                var element = data.ElementAt(i - 1);

                for (int j = 0; j < propertiesCount; j++)
                {
                    var property = properties.ElementAt(j);
                    CreateCell(row, j, element, property, workbook);
                }
            }

            sheet.SetAutoFilter(new CellRangeAddress(0, dataCount - 1, 0, propertiesCount - 1));
        }

        private static void CreateCell(IRow row, int columnIndex, T element, PropertyInfo property, IWorkbook workbook)
        {
            var value = property.GetValue(element);
            var format = workbook.CreateDataFormat();
            var cell = row.CreateCell(columnIndex);

            // Nie ma wbudowanego enuma na formatowanie i trzeba dopisać
            // Dla każdego typu trzeba rzutować, bo SetCellValue obsługuje tylko bool, string, double, IRichText, DateTime
            if (property.IsType<DateTime>())
            {
                var val = (DateTime)value;
                SetValueAndFormat(workbook, cell, val, 14);
            }
            else if (property.IsType<decimal>())
            {
                var val = Convert.ToDouble((decimal)value);
                SetValueAndFormat(workbook, cell, val, 2);
            }
            else
            {
                var val = value.ToString();
                cell.SetCellValue(val);

            }
        }

        //Nie ma wbudowanej obsługi nagłówków
        private static void CreateHeaderRow(IRow row, IEnumerable<PropertyInfo> properties)
        {
            var names = properties.Select(x => x.Name);

            for (int i = 0; i < names.Count(); i++)
            {
                var name = names.ElementAt(i);
                row.CreateCell(i).SetCellValue(names.ElementAt(i));
            }
        }

        private static void SetValueAndFormat(IWorkbook workbook, ICell cell, int value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }

        private static void SetValueAndFormat(IWorkbook workbook, ICell cell, double value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }

        private static void SetValueAndFormat(IWorkbook workbook, ICell cell, DateTime value, short formatId)
        {
            //set value for the cell
            if (value != null)
                cell.SetCellValue(value);

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }
    }
}
