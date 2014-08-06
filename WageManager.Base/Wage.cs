using System;

namespace WageManager
{
    namespace Base
    {
        public class Wage
        {
            private long Wageid;
            public long wageid
            {
                get { return Wageid; }
                set { Wageid = value; }
            }

            private long Employeeid;
            public long employeeid
            {
                get { return Employeeid; }
                set { Employeeid = value; }
            }

            private long Companyid;
            public long companyid
            {
                get { return Companyid; }
                set { Companyid = value; }
            }

            private long Companyid_tax;
            public long companyid_tax
            {
                get { return Companyid_tax; }
                set { Companyid_tax = value; }
            }

            private DateTime WageRound;
            public DateTime wageRound
            {
                get { return WageRound; }
                set { WageRound = value; }
            }

            private float BaseSalary;
            public float baseSalary
            {
                get { return BaseSalary; }
                set { BaseSalary = value; }
            }

            private float JobSalary;
            public float jobSalary
            {
                get { return JobSalary; }
                set { JobSalary = value; }
            }

            private float PerformanceBonus;
            public float performanceBonus
            {
                get { return PerformanceBonus; }
                set { PerformanceBonus = value; }
            }

            private float ProjectBonus;
            public float projectBonus
            {
                get { return ProjectBonus; }
                set { ProjectBonus = value; }
            }

            private float SaleBonus;
            public float saleBonus
            {
                get { return SaleBonus; }
                set { SaleBonus = value; }
            }

            private float AttendanceBonus;
            public float attendanceBonus
            {
                get { return AttendanceBonus; }
                set { AttendanceBonus = value; }
            }

            private float OvertimeBonus;
            public float overtimeBonus
            {
                get { return OvertimeBonus; }
                set { OvertimeBonus = value; }
            }

            private float AbsenceSalary;
            public float absenceSalary
            {
                get { return AbsenceSalary; }
                set { AbsenceSalary = value; }
            }

            private float AdjustmentSalary;
            public float adjustmentSalary
            {
                get { return AdjustmentSalary; }
                set { AdjustmentSalary = value; }
            }

            private float SocialWelfareDeduction;
            public float socialWelfareDeduction
            {
                get { return SocialWelfareDeduction; }
                set { SocialWelfareDeduction = value; }
            }

            private float PublicFundDeduction;
            public float publicFundDeduction
            {
                get { return PublicFundDeduction; }
                set { PublicFundDeduction = value; }
            }

            private float AdjustmentDeduction;
            public float adjustmentDeduction
            {
                get { return AdjustmentDeduction; }
                set { AdjustmentDeduction = value; }
            }

            private float MealBonus;
            public float mealBonus
            {
                get { return MealBonus; }
                set { MealBonus = value; }
            }

            private float HouseBonus;
            public float houseBonus
            {
                get { return HouseBonus; }
                set { HouseBonus = value; }
            }

            private float Allowance;
            public float allowance
            {
                get { return Allowance; }
                set { Allowance = value; }
            }
        }
    }
}