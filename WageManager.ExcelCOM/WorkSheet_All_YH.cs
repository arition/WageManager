using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WageManager.Base;

namespace WageManager.ExcelCOM
{
    class WorkSheet_All_YH
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            int currentRow = 10;
            int departmentStartRow = 11;
            bool departmentFlag = false;
            List<int> TotalWageList = new List<int>();
            string temp_department = "";
            ws.Cells[6, 3] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            foreach (Wage wage in WageList.Where((s) => s.company.公司名.Contains("优弧")))
            {
                if (temp_department != wage.employee.部门)
                {
                    //填充小计
                    if (departmentFlag)
                    {
                        ws.get_Range("B" + currentRow, "S" + currentRow).NumberFormat = "0.00";
                        ws.get_Range("A" + currentRow, "S" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                        ws.Cells[currentRow, 1] = "小计";
                        ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 3] = "=SUM(C" + departmentStartRow + ":C" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 4] = "=SUM(D" + departmentStartRow + ":D" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 5] = "=SUM(E" + departmentStartRow + ":E" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 6] = "=SUM(F" + departmentStartRow + ":F" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 7] = "=SUM(G" + departmentStartRow + ":G" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 8] = "=SUM(H" + departmentStartRow + ":H" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 9] = "=SUM(I" + departmentStartRow + ":I" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 10] = "=SUM(J" + departmentStartRow + ":J" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 11] = "=SUM(K" + departmentStartRow + ":K" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 12] = "=SUM(L" + departmentStartRow + ":L" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 13] = "=SUM(M" + departmentStartRow + ":M" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 14] = "=SUM(N" + departmentStartRow + ":N" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 15] = "=SUM(O" + departmentStartRow + ":O" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 16] = "=SUM(P" + departmentStartRow + ":P" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 17] = "=SUM(Q" + departmentStartRow + ":Q" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 18] = "=SUM(R" + departmentStartRow + ":R" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 19] = "=SUM(S" + departmentStartRow + ":S" + (currentRow - 1) + ")";
                        departmentStartRow = currentRow + 2;
                        TotalWageList.Add(currentRow);
                        currentRow++;
                    }
                    //填充部门
                    departmentFlag = true;
                    Range range = ws.get_Range("A" + currentRow, "S" + currentRow);
                    range.Merge(false);
                    ws.Cells[currentRow, 1] = wage.employee.部门;
                    currentRow++;
                    temp_department = wage.employee.部门;
                }
                //填充个人数据
                ws.get_Range("B" + currentRow, "S" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 1] = wage.employee.姓名;
                ws.Cells[currentRow, 2] = wage.baseSalary;
                ws.Cells[currentRow, 3] = wage.jobSalary;
                ws.Cells[currentRow, 4] = wage.performanceBonus;
                ws.Cells[currentRow, 5] = wage.houseBonus;
                ws.Cells[currentRow, 6] = wage.projectBonus;
                ws.Cells[currentRow, 7] = wage.saleBonus;
                ws.Cells[currentRow, 8] = wage.attendanceBonus;
                ws.Cells[currentRow, 9] = wage.overtimeBonus;
                ws.Cells[currentRow, 10] = wage.absenceSalary;
                ws.Cells[currentRow, 11] = wage.mealBonus;
                ws.Cells[currentRow, 12] = wage.adjustmentSalary;
                ws.Cells[currentRow, 13] = "=SUM(B" + currentRow + ":L" + currentRow + ")";
                ws.Cells[currentRow, 14] = wage.socialWelfareDeduction;
                ws.Cells[currentRow, 15] = wage.publicFundDeduction;
                ws.Cells[currentRow, 16] = Utils.CalcTax(wage);
                ws.Cells[currentRow, 17] = wage.adjustmentDeduction;
                ws.Cells[currentRow, 18] = 0;
                ws.Cells[currentRow, 19] = "=M" + currentRow + "-N" + currentRow + "-O" + currentRow + "-P" + currentRow + "-Q" + currentRow + "+R" + currentRow;
                currentRow++;
            }
            //填充最后一个小计
            if (departmentFlag)
            {
                ws.get_Range("B" + currentRow, "S" + currentRow).NumberFormat = "0.00";
                ws.get_Range("A" + currentRow, "S" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                ws.Cells[currentRow, 1] = "小计";
                ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 3] = "=SUM(C" + departmentStartRow + ":C" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 4] = "=SUM(D" + departmentStartRow + ":D" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 5] = "=SUM(E" + departmentStartRow + ":E" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 6] = "=SUM(F" + departmentStartRow + ":F" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 7] = "=SUM(G" + departmentStartRow + ":G" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 8] = "=SUM(H" + departmentStartRow + ":H" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 9] = "=SUM(I" + departmentStartRow + ":I" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 10] = "=SUM(J" + departmentStartRow + ":J" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 11] = "=SUM(K" + departmentStartRow + ":K" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 12] = "=SUM(L" + departmentStartRow + ":L" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 13] = "=SUM(M" + departmentStartRow + ":M" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 14] = "=SUM(N" + departmentStartRow + ":N" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 15] = "=SUM(O" + departmentStartRow + ":O" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 16] = "=SUM(P" + departmentStartRow + ":P" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 17] = "=SUM(Q" + departmentStartRow + ":Q" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 18] = "=SUM(R" + departmentStartRow + ":R" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 19] = "=SUM(S" + departmentStartRow + ":S" + (currentRow - 1) + ")";
                departmentStartRow = currentRow + 2;
                TotalWageList.Add(currentRow);
                currentRow++;
            }
            //填充合计
            ws.get_Range("B" + currentRow, "S" + currentRow).NumberFormat = "0.00";
            ws.get_Range("A" + currentRow, "S" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 204, 0));
            string allTotalWageCellAlgorithm = "=_Column_" + string.Join("+_Column_", TotalWageList.ToArray());
            ws.Cells[currentRow, 1] = "合计";
            ws.Cells[currentRow, 2] = allTotalWageCellAlgorithm.Replace("_Column_", "B");
            ws.Cells[currentRow, 3] = allTotalWageCellAlgorithm.Replace("_Column_", "C");
            ws.Cells[currentRow, 4] = allTotalWageCellAlgorithm.Replace("_Column_", "D");
            ws.Cells[currentRow, 5] = allTotalWageCellAlgorithm.Replace("_Column_", "E");
            ws.Cells[currentRow, 6] = allTotalWageCellAlgorithm.Replace("_Column_", "F");
            ws.Cells[currentRow, 7] = allTotalWageCellAlgorithm.Replace("_Column_", "G");
            ws.Cells[currentRow, 8] = allTotalWageCellAlgorithm.Replace("_Column_", "H");
            ws.Cells[currentRow, 9] = allTotalWageCellAlgorithm.Replace("_Column_", "I");
            ws.Cells[currentRow, 10] = allTotalWageCellAlgorithm.Replace("_Column_", "J");
            ws.Cells[currentRow, 11] = allTotalWageCellAlgorithm.Replace("_Column_", "K");
            ws.Cells[currentRow, 12] = allTotalWageCellAlgorithm.Replace("_Column_", "L");
            ws.Cells[currentRow, 13] = allTotalWageCellAlgorithm.Replace("_Column_", "M");
            ws.Cells[currentRow, 14] = allTotalWageCellAlgorithm.Replace("_Column_", "N");
            ws.Cells[currentRow, 15] = allTotalWageCellAlgorithm.Replace("_Column_", "O");
            ws.Cells[currentRow, 16] = allTotalWageCellAlgorithm.Replace("_Column_", "P");
            ws.Cells[currentRow, 17] = allTotalWageCellAlgorithm.Replace("_Column_", "Q");
            ws.Cells[currentRow, 18] = allTotalWageCellAlgorithm.Replace("_Column_", "R");
            ws.Cells[currentRow, 19] = allTotalWageCellAlgorithm.Replace("_Column_", "S");
            //设置边框
            ws.get_Range("A10", "S" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("A10", "S" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("A10", "S" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
        }
    }
}
