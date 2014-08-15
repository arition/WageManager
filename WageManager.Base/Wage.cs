using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel;

namespace WageManager
{
    namespace Base
    {
        public class Wage:INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            private void NotifyPropertyChanged(String propertyName = "")
            {
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                }
            }

            private long Wageid;
            public long wageid
            {
                get { return Wageid; }
                set { Wageid = value; }
            }

            private Employee Employee;
            public virtual Employee employee
            {
                get { return Employee; }
                set { Employee = value; }
            }

            private Company Company;
            public virtual Company company
            {
                get { return Company; }
                set { Company = value; }
            }

            private Company Company_tax;
            public virtual Company company_tax
            {
                get { return Company_tax; }
                set { Company_tax = value; }
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

            private float Overtime_weekDay;
            public float overtime_weekDay
            {
                get { return Overtime_weekDay; }
                set
                {
                    Overtime_weekDay = value;
                    try
                    {
                        overtimeBonus = System.Convert.ToSingle(Math.Round(System.Convert.ToDouble(overtime_weekDay * company.平时加班工资 + overtime_weekEnd * company.周末加班工资), 2));
                    }
                    catch { }
                    NotifyPropertyChanged("overtimeBonus");
                }
            }

            private float Overtime_weekEnd;
            public float overtime_weekEnd
            {
                get { return Overtime_weekEnd; }
                set
                {
                    Overtime_weekEnd = value;
                    try
                    {
                        overtimeBonus = System.Convert.ToSingle(Math.Round(System.Convert.ToDouble(overtime_weekDay * company.平时加班工资 + overtime_weekEnd * company.周末加班工资), 2));
                    }
                    catch { }
                    NotifyPropertyChanged("overtimeBonus");
                }
            }

            private float OvertimeBonus;
            public float overtimeBonus
            {
                get { return OvertimeBonus; }
                set { OvertimeBonus = value; }
            }

            private float AbsenceTime;
            public float absenceTime
            {
                get { return AbsenceTime; }
                set
                {
                    AbsenceTime = value;
                    absenceSalary = System.Convert.ToSingle(Math.Round(System.Convert.ToDouble(absenceTime * (baseSalary / 165f)), 2));
                    NotifyPropertyChanged("absenceSalary");
                }
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

            private float Tax;
            public float tax
            {
                get { return Tax; }
                set { Tax = value; }
            }
        }
    }
}