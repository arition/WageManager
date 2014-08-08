using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using WageManager.Base;
using System.Drawing;

namespace WageManager
{
    namespace ExcelCOM
    {
        public static class CreateExcel
        {
            private class WageComparer : IComparer<Wage>
            {
                public int Compare(Wage x, Wage y)
                {
                    return x.employee.部门.CompareTo(y.employee.部门);
                }
            }

            public static void Create(List<Wage> WageList)
            {
                Application xlApp = new Application();
                if (xlApp == null)
                {
                    throw new InvalidOperationException("EXCEL could not be started.");
                }
                xlApp.Visible = true;
                Workbook wb = xlApp.Workbooks.Open(@"D:\My Program\WageManager\WageManager\bin\Debug\DCMC_Pay_FT.xls", Editable: true, IgnoreReadOnlyRecommended: true);

                WageList.Sort(new WageComparer());

                Worksheet ws1 = (Worksheet)wb.Worksheets[1];
                WorkSheet_First.Create(ws1, WageList);

                Worksheet ws2 = (Worksheet)wb.Worksheets[2];
                WorkSheet_First.Create(ws2, WageList);
                
            }
        }
    }
}