namespace WageManager
{
    namespace Base
    {
        public class Company
        {
            private long Companyid;
            public long companyid
            {
                get { return Companyid; }
                set { Companyid = value; }
            }

            private string Name;
            public string 公司名
            {
                get { return Name; }
                set { Name = value; }
            }

            private float OvertimeSalary_Weekday;
            public float 平时加班工资
            {
                get { return OvertimeSalary_Weekday; }
                set { OvertimeSalary_Weekday = value; }
            }

            private float OvertimeSalary_Weekend;
            public float 周末加班工资
            {
                get { return OvertimeSalary_Weekend; }
                set { OvertimeSalary_Weekend = value; }
            }
        }
    }
}