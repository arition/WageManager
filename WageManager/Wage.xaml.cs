using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WageManager.Database;
using WageManager.Base;
using System.Data.Entity;
using MahApps.Metro.Controls;

namespace WageManager
{
    /// <summary>
    /// Wage.xaml 的交互逻辑
    /// </summary>
    public partial class Wage : MetroWindow
    {
        public Wage()
        {
            InitializeComponent();
        }

        DatabaseContext db = new DatabaseContext();
        

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            //IQueryable<Base.Wage> lastMonthWage = db.Wages.Where(o => (o.wageRound.Year == DateTime.Now.Year) && (DateTime.Now.Month - o.wageRound.Month == 2));
            //var ww = db.Employees;
        }
    }
}
