namespace WageManager
{
    namespace Base
    {
        public class Company
        {
            private int Id;
            public int id
            {
                get { return Id; }
                set { Id = value; }
            }

            private string Name;
            public string name
            {
                get { return Name; }
                set { Name = value; }
            }

            private float OvertimeSalary_Weekday;
            public float overtimeSalary_Weekday
            {
                get { return OvertimeSalary_Weekday; }
                set { OvertimeSalary_Weekday = value; }
            }

            private float OvertimeSalary_Weekend;
            public float overtimeSalary_Weekend
            {
                get { return OvertimeSalary_Weekend; }
                set { OvertimeSalary_Weekend = value; }
            }
            

        }
    }
}