using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WageManager.Base;

namespace WageManager.ExcelCOM
{
    class WorkSheet_Tax
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            int currentRow = 9;
            ws.Cells[5, 3] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            foreach (Wage wage in WageList)
            {
                ws.get_Range("C" + currentRow, "J" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 1] = wage.employee.姓名;
                ws.Cells[currentRow, 2] = wage.employee.身份证号;
                ws.Cells[currentRow, 3] = 
                    wage.baseSalary + wage.jobSalary + wage.performanceBonus + wage.projectBonus +
                    wage.saleBonus + wage.attendanceBonus + wage.overtimeBonus + wage.absenceSalary +
                    wage.adjustmentSalary;
                ws.Cells[currentRow, 4] = wage.socialWelfareDeduction + wage.publicFundDeduction;
                ws.Cells[currentRow, 5] = 3500;
                ws.Cells[currentRow, 6] = "=C" + currentRow + "-D" + currentRow + "-E" + currentRow;
                ws.Cells[currentRow, 7] = "=IF(F" + currentRow + "<0,0,IF(F" + currentRow + "<1500,F" + currentRow + "*3%,IF(F" + currentRow + "<4500,F" + currentRow + "*10%-105,IF(F" + currentRow + "<9000,F" + currentRow + "*20%-555,IF(F" + currentRow + "<35000,F" + currentRow + "*25%-1005,IF(F" + currentRow + "<55000,F" + currentRow + "*30%-2275,))))))";
                ws.Cells[currentRow, 8] = wage.adjustmentDeduction;
                ws.Cells[currentRow, 9] = "=C" + currentRow + "-D" + currentRow + "-G" + currentRow + "-H" + currentRow;
                ws.Cells[currentRow, 10] = "=C" + currentRow + "-D" + currentRow;
                ws.Cells[currentRow, 11] = wage.company_tax.公司名;
                currentRow++;
            }
            ws.get_Range("C" + currentRow, "J" + currentRow).NumberFormat = "0.00";
            ws.get_Range("A" + currentRow, "K" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 204, 0));
            ws.Cells[currentRow, 1] = "合计";
            ws.Cells[currentRow, 3] = "=SUM(C" + 9 + ":C" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 4] = "=SUM(D" + 9 + ":D" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 5] = "=SUM(E" + 9 + ":E" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 6] = "=SUM(F" + 9 + ":F" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 7] = "=SUM(G" + 9 + ":G" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 8] = "=SUM(H" + 9 + ":H" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 9] = "=SUM(I" + 9 + ":I" + (currentRow - 1) + ")";
            ws.Cells[currentRow, 10] = "=SUM(J" + 9 + ":J" + (currentRow - 1) + ")";
            //设置边框
            ws.get_Range("A9", "K" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("A9", "K" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("A9", "K" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
        }
    }
}
