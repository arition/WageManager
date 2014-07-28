namespace WageManager
{
    namespace Base
    {
        public class Employee
        {
            private long Id;
            public long id
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

            private string Department;

            public string department
            {
                get { return Department; }
                set { Department = value; }
            }

            private float LastBaseSalary;

            public float lastBaseSalary
            {
                get { return LastBaseSalary; }
                set { LastBaseSalary = value; }
            }

            private string BankCardNumber;

            public string bankCardNumber
            {
                get { return BankCardNumber; }
                set { BankCardNumber = value; }
            }

            private string Bank;

            public string bank
            {
                get { return Bank; }
                set { Bank = value; }
            }

            private string IdCardNumber;

            public string idCardNumber
            {
                get { return IdCardNumber; }
                set { IdCardNumber = value; }
            }
            
        }
    }
}