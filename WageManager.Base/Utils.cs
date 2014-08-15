using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WageManager.Base
{
    public static class Utils
    {
        public static float CalcTax(Wage wage)
        {
            float tax = 0;
            if (wage.tax != 0)
            {
                tax = wage.tax;
            }
            else
            {
                float base_Salary =
                        wage.baseSalary + wage.jobSalary + wage.performanceBonus + wage.projectBonus +
                        wage.saleBonus + wage.attendanceBonus + wage.overtimeBonus + wage.absenceSalary +
                        wage.adjustmentSalary - (wage.socialWelfareDeduction + wage.publicFundDeduction) - 3500;
                if (base_Salary < 0) { tax = 0; }
                else if (base_Salary < 1500) { tax = base_Salary * 0.03f; }
                else if (base_Salary < 4500) { tax = base_Salary * 0.1f - 105; }
                else if (base_Salary < 9000) { tax = base_Salary * 0.2f - 555; }
                else if (base_Salary < 35000) { tax = base_Salary * 0.25f - 1005; }
                else if (base_Salary < 55000) { tax = base_Salary * 0.3f - 2275; }
            }
            return tax;
        }
    }
}
