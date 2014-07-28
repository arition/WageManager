using System;

namespace WageManager
{
    namespace Base
    {
        public class Wage
        {
            private long Id;
            public long id
            {
                get { return Id; }
                set { Id = value; }
            }

            private long Userid;
            public long userid
            {
                get { return Userid; }
                set { Userid = value; }
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

            private float Allowance;
            public float allowance
            {
                get { return Allowance; }
                set { Allowance = value; }
            }
            
        }
    }
}