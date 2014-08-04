using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WageManager.Database;
using WageManager.Base;
using System.Data.Entity;
using MahApps.Metro.Controls;

namespace WageManager
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //var db = new DatabaseContext();
            //Database.Initializer.Initialize(db);
            //var result = db.Companys.OrderBy(c => c.name).ToList();
            //MessageBox.Show(result[0].name);
        }

        private void btn_Employee_Click(object sender, RoutedEventArgs e)
        {
            var window_Employee = new Employee();
            window_Employee.Show();
        }

        private void btn_Company_Click(object sender, RoutedEventArgs e)
        {
            var window_Company = new Company();
            window_Company.Show();
        }
    }
}
