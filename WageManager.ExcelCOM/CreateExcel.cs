using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WageManager.Base;

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
                xlApp.Visible = false;
                Workbook wb = xlApp.Workbooks.Open(@"D:\My Program\WageManager\WageManager\bin\Debug\DCMC_Pay_FT.xls", Editable: true, IgnoreReadOnlyRecommended: true);

                WageList.Sort(new WageComparer());

                Worksheet ws1 = (Worksheet)wb.Worksheets[1];
                WorkSheet_All_DWDC.Create(ws1, WageList);

                Worksheet ws2 = (Worksheet)wb.Worksheets[2];
                WorkSheet_All_YH.Create(ws2, WageList);

                Worksheet ws3 = (Worksheet)wb.Worksheets[3];
                WorkSheet_Overtime.Create(ws3, WageList);

                Worksheet ws4 = (Worksheet)wb.Worksheets[4];
                WorkSheet_Absence.Create(ws4, WageList);

                Worksheet ws5 = (Worksheet)wb.Worksheets[5];
                WorkSheet_ProjectBonus.Create(ws5, WageList);

                Worksheet ws6 = (Worksheet)wb.Worksheets[6];
                WorkSheet_AllSalary_DWDC.Create(ws6, WageList);

                Worksheet ws7 = (Worksheet)wb.Worksheets[7];
                WorkSheet_AllSalary_YH.Create(ws7, WageList);

                Worksheet ws8 = (Worksheet)wb.Worksheets[8];
                WorkSheet_Tax.Create(ws8, WageList);

                Worksheet ws9 = (Worksheet)wb.Worksheets[9];
                WorkSheet_Individual_Wage_DWDC.Create(ws9, WageList);

                Worksheet ws10 = (Worksheet)wb.Worksheets[10];
                WorkSheet_Individual_Wage_YH.Create(ws10, WageList);

                xlApp.Visible = true;

                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}