using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using OfficeOpenXml;
using System.IO;

namespace Excel_Reader
{
    internal class ExcelManager
    {
        string fileLocation = Path.Combine(Directory.GetCurrentDirectory(), ConfigurationManager.AppSettings.Get("ExcelFileName"));
        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            //first row for properties
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n => new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() });

            for (int row = 2; row <= sheet.Dimension.Rows; row++)
            { 
                T obj = (T)Activator.CreateInstance(typeof(T)); //generic object
                foreach (var prop in typeof(T).GetProperties())
                {
                    //int col = sheet.Dimension.Columns;
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);

            }
            return list;
        }

        internal List<Models.ExcelSheetData> GetExcelData()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileLocation)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets[0];
                var acquisitions = new ExcelManager().GetList<Models.ExcelSheetData>(sheet);
                return acquisitions;
            }
        }
    }
}
