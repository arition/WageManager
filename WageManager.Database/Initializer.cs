using System.Collections.Generic;
using System.Data.Entity;
using WageManager.Base;
using System.IO;
namespace WageManager
{
    namespace Database
    {
        public class Initializer
        {
            public static void Initialize(DatabaseContext context)
            {
                if (File.Exists(Path.Combine("data", "data.s3db"))) { File.Delete(Path.Combine("data", "data.s3db")); }
                File.Copy(Path.Combine("data", "data_backup.s3db"), Path.Combine("data", "data.s3db"));
                var companys = new List<Company>
                {
                    new Company{公司名="典驰", 平时加班工资=15, 周末加班工资=15},
                    new Company{公司名="优弧", 平时加班工资=13.5f, 周末加班工资=32}
                };
                var employees = new List<Employee>
                {
                    new Employee{姓名="Testa",部门="销售部",基础工资=1500},
                    new Employee{姓名="Testb",部门="销售部",基础工资=1600},
                    new Employee{姓名="Testc",部门="仓储部",基础工资=1500},
                    new Employee{姓名="Testd",部门="管理部",基础工资=2000}
                };
                companys.ForEach(s => context.Companys.Add(s));
                employees.ForEach(s => context.Employees.Add(s));
                context.SaveChanges();
                var wages = new List<Wage>
                {
                    new Wage(){
                        employee=context.Employees.Find(1),
                        company=context.Companys.Find(1),
                        company_tax=context.Companys.Find(1),
                        wageRound=new System.DateTime(2014,6,1),
                        baseSalary=1500,
                        jobSalary=100,
                        performanceBonus=100,
                        projectBonus=0,
                        saleBonus=122.5f,
                        attendanceBonus=10,
                        overtimeBonus=5,
                        absenceSalary=10,
                        adjustmentSalary=0,
                        socialWelfareDeduction=12.56f,
                        publicFundDeduction=13.01f,
                        adjustmentDeduction=0,
                        allowance=0
                    },
                    new Wage(){
                        employee=context.Employees.Find(2),
                        company=context.Companys.Find(1),
                        company_tax=context.Companys.Find(2),
                        wageRound=new System.DateTime(2014,6,1),
                        baseSalary=2000,
                        jobSalary=100,
                        performanceBonus=100,
                        projectBonus=0,
                        saleBonus=122.5f,
                        attendanceBonus=10,
                        overtimeBonus=5,
                        absenceSalary=10,
                        adjustmentSalary=0,
                        socialWelfareDeduction=12.56f,
                        publicFundDeduction=13.01f,
                        adjustmentDeduction=0,
                        allowance=0
                    },
                    new Wage(){
                        employee=context.Employees.Find(3),
                        company=context.Companys.Find(1),
                        company_tax=context.Companys.Find(2),
                        wageRound=new System.DateTime(2014,6,1),
                        baseSalary=2000,
                        jobSalary=500,
                        performanceBonus=200,
                        projectBonus=0,
                        saleBonus=12.5f,
                        attendanceBonus=130,
                        overtimeBonus=25,
                        absenceSalary=110,
                        adjustmentSalary=0,
                        socialWelfareDeduction=14.56f,
                        publicFundDeduction=16.01f,
                        adjustmentDeduction=0,
                        allowance=0
                    },
                    new Wage(){
                        employee=context.Employees.Find(4),
                        company=context.Companys.Find(1),
                        company_tax=context.Companys.Find(2),
                        wageRound=new System.DateTime(2014,6,1),
                        baseSalary=2000,
                        jobSalary=500,
                        performanceBonus=200,
                        projectBonus=0,
                        saleBonus=12.5f,
                        attendanceBonus=130,
                        overtimeBonus=25,
                        absenceSalary=110,
                        adjustmentSalary=0,
                        socialWelfareDeduction=14.56f,
                        publicFundDeduction=16.01f,
                        adjustmentDeduction=0,
                        allowance=0
                    }
                };
                wages.ForEach(s => context.Wages.Add(s));
                context.SaveChanges();
            }
        }
    }
}
