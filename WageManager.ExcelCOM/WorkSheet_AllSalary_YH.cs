using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using WageManager.Base;
using System.Drawing;

namespace WageManager.ExcelCOM
{
    class WorkSheet_AllSalary_YH
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            //迪典填充
            int currentRow = 8;
            string temp_department = "";
            ws.Cells[4, 4] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            foreach (Wage wage in WageList.Where((s) => s.company.公司名.Contains("优弧")))
            {
                if (temp_department != wage.employee.部门)
                {
                    //填充部门
                    Range range = ws.get_Range("A" + currentRow, "D" + currentRow);
                    range.Merge(false);
                    ws.Cells[currentRow, 1] = wage.employee.部门;
                    currentRow++;
                    temp_department = wage.employee.部门;
                }
                //填充个人数据
                ws.get_Range("D" + currentRow, "D" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 1] = wage.employee.姓名;
                ws.Cells[currentRow, 2] = wage.employee.银行卡号;
                ws.Cells[currentRow, 3] = wage.employee.开户行;
                ws.Cells[currentRow, 4] =
                    wage.baseSalary + wage.jobSalary + wage.performanceBonus + wage.projectBonus +
                    wage.saleBonus + wage.attendanceBonus + wage.overtimeBonus + wage.absenceSalary +
                    wage.adjustmentSalary - wage.socialWelfareDeduction - wage.publicFundDeduction -
                    wage.adjustmentDeduction + wage.mealBonus;
                currentRow++;
            }
            currentRow--;
            //设置边框
            ws.get_Range("A8", "D" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("A8", "D" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("A8", "D" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
        }
    }
}
