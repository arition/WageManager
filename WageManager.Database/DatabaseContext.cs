using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using WageManager.Base;

namespace WageManager
{
    namespace Database
    {
        public class DatabaseContext : DbContext
        {
            public DbSet<Employee> Employees { get; set; }
            public DbSet<Company> Companys { get; set; }
            public DbSet<Wage> Wages { get; set; }
            protected override void OnModelCreating(DbModelBuilder modelBuilder)
            {
                modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            }
        }
    }
}
