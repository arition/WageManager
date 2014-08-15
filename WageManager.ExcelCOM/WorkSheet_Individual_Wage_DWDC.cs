using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WageManager.Base;

namespace WageManager.ExcelCOM
{
    class WorkSheet_Individual_Wage_DWDC
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            int currentRow = 1;
            foreach (Wage wage in WageList.Where((s) => !s.company.公司名.Contains("优弧")))
            {
                string DeductionEnd_str = "M";
                int Wage_int = 14;
                string Wage_str = "N";

                if (wage.socialWelfareDeduction != 0)
                {
                    DeductionEnd_str = "N";
                    Wage_int++;
                    Wage_str = "O";
                    if (wage.publicFundDeduction != 0)
                    {
                        DeductionEnd_str = "O";
                        Wage_int++;
                        Wage_str = "P";
                    }
                }
                if (wage.socialWelfareDeduction == 0 && wage.publicFundDeduction != 0)
                {
                    DeductionEnd_str = "N";
                    Wage_int++;
                    Wage_str = "O";
                }

                Range range = ws.get_Range("A" + currentRow, Wage_str + currentRow);
                range.Merge(false);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[currentRow, 1] = "员工工资表";
                currentRow += 2;
                range = ws.get_Range("A" + currentRow, Wage_str + currentRow);
                range.Merge(false);
                ws.Cells[currentRow, 1] = "单位："+wage.company.公司名;

                currentRow++;
                ws.Cells[currentRow, 1] = "薪金月份";
                ws.Cells[currentRow, 3] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
                ws.get_Range("C" + currentRow, "C" + currentRow).NumberFormat = "yyyy年mm月";

                currentRow += 2;
                range = ws.get_Range("A" + currentRow, "A" + (currentRow+1));
                range.Merge(false);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                ws.Cells[currentRow, 1] = "姓  名";
                range = ws.get_Range("B" + currentRow, "K" + currentRow);
                range.Merge(false);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[currentRow, 2] = "应发工资";
                range = ws.get_Range("L" + currentRow, DeductionEnd_str + currentRow);
                range.Merge(false);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[currentRow, 12] = "代扣费用";
                range = ws.get_Range(Wage_str + currentRow, Wage_str + (currentRow + 1));
                range.Merge(false);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                ws.Cells[currentRow, Wage_int] = "实领金额";

                currentRow++;
                int currentColumn = 2;
                ws.Cells[currentRow, currentColumn] = "基本薪金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "岗位薪金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "绩效奖金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "项目奖金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "销售奖金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "全勤奖金"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "加班费"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "缺勤金额"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "调整费用"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "应付薪资"; currentColumn++;
                if (wage.socialWelfareDeduction != 0) { ws.Cells[currentRow, currentColumn] = "社保费"; currentColumn++; }
                if (wage.publicFundDeduction != 0) { ws.Cells[currentRow, currentColumn] = "公积金"; currentColumn++; }
                ws.Cells[currentRow, currentColumn] = "所得税"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "调整费用"; currentColumn++;

                currentRow++;
                currentColumn = 1;
                ws.get_Range("B" + currentRow, Wage_str + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, currentColumn] = wage.employee.姓名; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.baseSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.jobSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.performanceBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.projectBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.saleBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.attendanceBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.overtimeBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.absenceSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.adjustmentSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "=SUM(B" + currentRow + ":J" + currentRow + ")"; currentColumn++;
                if (wage.socialWelfareDeduction != 0) { ws.Cells[currentRow, currentColumn] = wage.socialWelfareDeduction; currentColumn++; }
                if (wage.publicFundDeduction != 0) { ws.Cells[currentRow, currentColumn] = wage.publicFundDeduction; currentColumn++; }
                ws.Cells[currentRow, currentColumn] = Utils.CalcTax(wage); currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.adjustmentDeduction; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "=SUM(B" + currentRow + ":K" + currentRow + ")-SUM(L" + currentRow + ":" + DeductionEnd_str + currentRow + ")"; currentColumn++;

                currentRow++;
                currentColumn = 1;
                ws.get_Range("B" + currentRow, Wage_str + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, currentColumn] = "合计"; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.baseSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.jobSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.performanceBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.projectBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.saleBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.attendanceBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.overtimeBonus; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.absenceSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.adjustmentSalary; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "=SUM(B" + currentRow + ":J" + currentRow + ")"; currentColumn++;
                if (wage.socialWelfareDeduction != 0) { ws.Cells[currentRow, currentColumn] = wage.socialWelfareDeduction; currentColumn++; }
                if (wage.publicFundDeduction != 0) { ws.Cells[currentRow, currentColumn] = wage.publicFundDeduction; currentColumn++; }
                ws.Cells[currentRow, currentColumn] = Utils.CalcTax(wage); currentColumn++;
                ws.Cells[currentRow, currentColumn] = wage.adjustmentDeduction; currentColumn++;
                ws.Cells[currentRow, currentColumn] = "=SUM(B" + currentRow + ":K" + currentRow + ")-SUM(L" + currentRow + ":" + DeductionEnd_str + currentRow + ")"; currentColumn++;

                //设置边框
                ws.get_Range("A" + (currentRow - 3), Wage_str + currentRow).Borders.Weight = XlBorderWeight.xlThin;
                ws.get_Range("A" + (currentRow - 3), Wage_str + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
                ws.get_Range("A" + (currentRow - 3), Wage_str + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
                currentRow += 4;
            }
            ws.Cells.Columns.AutoFit();
        }
    }
}
