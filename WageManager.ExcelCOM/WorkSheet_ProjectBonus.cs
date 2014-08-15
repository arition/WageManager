using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using WageManager.Base;

namespace WageManager.ExcelCOM
{
    class WorkSheet_ProjectBonus
    {
        public static void Create(Worksheet ws, List<Wage> WageList)
        {
            int currentRow = 7;
            int departmentStartRow = 8;
            bool departmentFlag = false;
            List<int> TotalWageList = new List<int>();
            string temp_department = "";
            ws.Cells[4, 2] = DateTime.Now.Year + "年" + (DateTime.Now.Month - 1) + "月";
            foreach (Wage wage in WageList)
            {
                if (temp_department != wage.employee.部门)
                {
                    //填充小计
                    if (departmentFlag)
                    {
                        ws.get_Range("B" + currentRow, "B" + currentRow).NumberFormat = "0.00";
                        ws.get_Range("A" + currentRow, "B" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                        ws.Cells[currentRow, 1] = "小计";
                        ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                        departmentStartRow = currentRow + 2;
                        TotalWageList.Add(currentRow);
                        currentRow++;
                    }
                    //填充部门
                    departmentFlag = true;
                    Range range = ws.get_Range("A" + currentRow, "B" + currentRow);
                    range.Merge(false);
                    ws.Cells[currentRow, 1] = wage.employee.部门;
                    currentRow++;
                    temp_department = wage.employee.部门;
                }
                //填充个人数据
                ws.get_Range("B" + currentRow, "B" + currentRow).NumberFormat = "0.00";
                ws.Cells[currentRow, 1] = wage.employee.姓名;
                ws.Cells[currentRow, 2] = wage.projectBonus;
                currentRow++;
            }
            //填充最后一个小计
            if (departmentFlag)
            {
                ws.get_Range("B" + currentRow, "B" + currentRow).NumberFormat = "0.00";
                ws.get_Range("A" + currentRow, "B" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                ws.Cells[currentRow, 1] = "小计";
                ws.Cells[currentRow, 2] = "=SUM(B" + departmentStartRow + ":B" + (currentRow - 1) + ")";
                departmentStartRow = currentRow + 2;
                TotalWageList.Add(currentRow);
                currentRow++;
            }
            //填充合计
            ws.get_Range("B" + currentRow, "B" + currentRow).NumberFormat = "0.00";
            ws.get_Range("A" + currentRow, "B" + currentRow).Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 204, 0));
            string allTotalWageCellAlgorithm = "=_Column_" + string.Join("+_Column_", TotalWageList.ToArray());
            ws.Cells[currentRow, 1] = "合计";
            ws.Cells[currentRow, 2] = allTotalWageCellAlgorithm.Replace("_Column_", "B");
            //设置边框
            ws.get_Range("A7", "B" + currentRow).Borders.Weight = XlBorderWeight.xlThin;
            ws.get_Range("A7", "B" + currentRow).Borders.LineStyle = XlLineStyle.xlContinuous;
            ws.get_Range("A7", "B" + currentRow).BorderAround2(LineStyle: XlLineStyle.xlContinuous, Weight: XlBorderWeight.xlMedium);
        }
    }
}
