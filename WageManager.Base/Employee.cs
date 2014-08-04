namespace WageManager
{
    namespace Base
    {
        public class Employee
        {
            private long Employeeid;
            public long employeeid
            {
                get { return Employeeid; }
                set { Employeeid = value; }
            }

            private string Name;
            public string 姓名
            {
                get { return Name; }
                set { Name = value; }
            }

            private string Department;
            public string 部门
            {
                get { return Department; }
                set { Department = value; }
            }

            private float LastBaseSalary;
            public float 基础工资
            {
                get { return LastBaseSalary; }
                set { LastBaseSalary = value; }
            }

            private string BankCardNumber;
            public string 银行卡号
            {
                get { return BankCardNumber; }
                set { BankCardNumber = value; }
            }

            private string Bank;
            public string 开户行
            {
                get { return Bank; }
                set { Bank = value; }
            }

            private string IdCardNumber;
            public string 身份证号
            {
                get { return IdCardNumber; }
                set { IdCardNumber = value; }
            }
            
        }
    }
}