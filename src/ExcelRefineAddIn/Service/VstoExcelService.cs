using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefineAddIn.Service
{
    public class VstoExcelService
    {
        private static readonly Lazy<VstoExcelService> _instance = new Lazy<VstoExcelService>(() => new VstoExcelService());

        public static VstoExcelService Instance => _instance.Value;

        private VstoExcelService() {}

        public List<List<object>> ReadActiveSheet(Worksheet sheet)
        {
            Range usedRange = null;
            var rows = new List<List<object>>();

            try
            {
                usedRange = sheet.UsedRange;

                for (int row = 1; row <= usedRange.Rows.Count; row++)
                {
                    var rowData = new List<object>();
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        Range cell = usedRange.Cells[row, col];
                        object value = cell.Value2;
                        rowData.Add(value);
                        Marshal.ReleaseComObject(cell);
                    }
                    rows.Add(rowData);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if(usedRange != null)
                {
                    Marshal.FinalReleaseComObject(usedRange);
                }
            }
            return rows;
        }
    }
}
