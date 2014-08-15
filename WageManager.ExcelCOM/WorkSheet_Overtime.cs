using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WageManager.Base;

namespace WageManager.ExcelCOM
{
    class WorkSheet_Overtime
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            //迪典填充
            int currentRow = 13;
            int departmentStartRow = 14;
            bool departmentFlag = false;
            List<int> TotalWageList = new List<int>();
            string temp_department = "";
            ws.Cells[7, 2] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            ws.Cells[8, 4] = WageList.Where((s) => !s.company.公司名.Contains("优弧")).FirstOrDefault().company.平时加班工资;
            ws.Cells[9, 4] = WageList.Where((s) => !s.company.公司名.Contains("优弧")).FirstOrDefault().company.周末加班工资;
            ws.get_Range("D8" , "D9").NumberFormat = "0.00";
            foreach (Wage wage in WageList.Where((s) => !s.company.公司名.Contains("优弧")))
            {
                if (temp_department != wage.employee.部门)
                {
                    //填充小计
                    if (departmentFlag)
                    {
                        ws.get_Range("B" + currentRow, "D" + currentRow).NumberFormat = "0.00";
                        ws.get_Range("A" + currentRow, "D" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                        ws.Cells[currentRow, 1] = "小计";
                        ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 3] = "=SUM(C" + departmentStartRow + ":C" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 4] = "=SUM(D" + departmentStartRow + ":D" + (currentRow - 1) + ")";
                        departmentStartRow = currentRow + 2;
                        TotalWageList.Add(currentRow);
                        currentRow++;
                    }
                    //填充部门
                    departmentFlag = true;
                    Range range = ws.get_Range("A" + currentRow, "D" + currentRow);
                    range.Merge(false);
                    ws.Cells[currentRow, 1] = wage.employee.部门;
                    currentRow++;
                    temp_department = wage.employee.部门;
                }
                //填充个人数据
                ws.get_Range("B" + currentRow, "D" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 1] = wage.employee.姓名;
                ws.Cells[currentRow, 2] = wage.overtime_weekDay;
                ws.Cells[currentRow, 3] = wage.overtime_weekEnd;
                ws.Cells[currentRow, 4] = wage.overtimeBonus;
                currentRow++;
            }
            //填充最后一个小计
            if (departmentFlag)
            {
                ws.get_Range("B" + currentRow, "D" + currentRow).NumberFormat = "0.00";
                ws.get_Range("A" + currentRow, "D" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                ws.Cells[currentRow, 1] = "小计";
                ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 3] = "=SUM(C" + departmentStartRow + ":C" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 4] = "=SUM(D" + departmentStartRow + ":D" + (currentRow - 1) + ")";
                departmentStartRow = currentRow + 2;
                TotalWageList.Add(currentRow);
                currentRow++;
            }
            //填充合计
            ws.get_Range("B" + currentRow, "D" + currentRow).NumberFormat = "0.00";
            ws.get_Range("A" + currentRow, "D" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 204, 0));
            string allTotalWageCellAlgorithm = "=_Column_" + string.Join("+_Column_", TotalWageList.ToArray());
            ws.Cells[currentRow, 1] = "合计";
            ws.Cells[currentRow, 2] = allTotalWageCellAlgorithm.Replace("_Column_", "B");
            ws.Cells[currentRow, 3] = allTotalWageCellAlgorithm.Replace("_Column_", "C");
            ws.Cells[currentRow, 4] = allTotalWageCellAlgorithm.Replace("_Column_", "D");
            //设置边框
            ws.get_Range("A13", "D" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("A13", "D" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("A13", "D" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);


            //优弧填充
            currentRow = 13;
            departmentStartRow = 14;
            departmentFlag = false;
            TotalWageList = new List<int>();
            temp_department = "";
            ws.Cells[7, 10] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            ws.Cells[8, 11] = WageList.Where((s) => s.company.公司名.Contains("优弧")).FirstOrDefault().company.平时加班工资;
            ws.Cells[9, 11] = WageList.Where((s) => s.company.公司名.Contains("优弧")).FirstOrDefault().company.周末加班工资;
            ws.get_Range("K8", "K9").NumberFormat = "0.00";
            foreach (Wage wage in WageList.Where((s) => s.company.公司名.Contains("优弧")))
            {
                if (temp_department != wage.employee.部门)
                {
                    //填充小计
                    if (departmentFlag)
                    {
                        ws.get_Range("J" + currentRow, "L" + currentRow).NumberFormat = "0.00";
                        ws.get_Range("I" + currentRow, "L" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                        ws.Cells[currentRow, 9] = "小计";
                        ws.Cells[currentRow, 10] = "=SUM(J" + departmentStartRow + ":J" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 11] = "=SUM(K" + departmentStartRow + ":K" + (currentRow - 1) + ")";
                        ws.Cells[currentRow, 12] = "=SUM(L" + departmentStartRow + ":L" + (currentRow - 1) + ")";
                        departmentStartRow = currentRow + 2;
                        TotalWageList.Add(currentRow);
                        currentRow++;
                    }
                    //填充部门
                    departmentFlag = true;
                    Range range = ws.get_Range("I" + currentRow, "L" + currentRow);
                    range.Merge(false);
                    ws.Cells[currentRow, 9] = wage.employee.部门;
                    currentRow++;
                    temp_department = wage.employee.部门;
                }
                //填充个人数据
                ws.get_Range("J" + currentRow, "L" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 9] = wage.employee.姓名;
                ws.Cells[currentRow, 10] = wage.overtime_weekDay;
                ws.Cells[currentRow, 11] = wage.overtime_weekEnd;
                ws.Cells[currentRow, 12] = wage.overtimeBonus;
                currentRow++;
            }
            //填充最后一个小计
            if (departmentFlag)
            {
                ws.get_Range("J" + currentRow, "L" + currentRow).NumberFormat = "0.00";
                ws.get_Range("I" + currentRow, "L" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                ws.Cells[currentRow, 9] = "小计";
                ws.Cells[currentRow, 10] = "=SUM(J" + departmentStartRow + ":J" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 11] = "=SUM(K" + departmentStartRow + ":K" + (currentRow - 1) + ")";
                ws.Cells[currentRow, 12] = "=SUM(L" + departmentStartRow + ":L" + (currentRow - 1) + ")";
                departmentStartRow = currentRow + 2;
                TotalWageList.Add(currentRow);
                currentRow++;
            }
            //填充合计
            ws.get_Range("J" + currentRow, "L" + currentRow).NumberFormat = "0.00";
            ws.get_Range("I" + currentRow, "L" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 204, 0));
            allTotalWageCellAlgorithm = "=_Column_" + string.Join("+_Column_", TotalWageList.ToArray());
            ws.Cells[currentRow, 9] = "合计";
            ws.Cells[currentRow, 10] = allTotalWageCellAlgorithm.Replace("_Column_", "J");
            ws.Cells[currentRow, 11] = allTotalWageCellAlgorithm.Replace("_Column_", "K");
            ws.Cells[currentRow, 12] = allTotalWageCellAlgorithm.Replace("_Column_", "L");
            //设置边框
            ws.get_Range("I13", "L" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("I13", "L" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("I13", "L" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
        }
    }
}
