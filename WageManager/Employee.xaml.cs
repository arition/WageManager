using MahApps.Metro.Controls;
using System;
using System.Collections.ObjectModel;
using System.Data.Entity;
using System.Threading.Tasks;
using System.Windows;
using WageManager.Database;

namespace WageManager
{
    /// <summary>
    /// Employee.xaml 的交互逻辑
    /// </summary>
    public partial class Employee : MetroWindow
    {
        public Employee()
        {
            InitializeComponent();
        }

        DatabaseContext db = new DatabaseContext();
        ObservableCollection<Base.Employee> data;

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                db.Employees.Load();
                data = db.Employees.Local;
                this.Dispatcher.Invoke(new Action(() =>
                {
                    DataGrid_employee.ItemsSource = data;
                    DataGrid_employee.Visibility = Visibility.Visible;
                    ProgressRing_employee.IsActive = false;
                }));
            });
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ProgressRing_employee.IsActive = true;
            Task.Factory.StartNew(() =>
            {
                db.SaveChanges();
            });
        }

        private void MenuItem_delete_Click(object sender, RoutedEventArgs e)
        {
            data.RemoveAt(DataGrid_employee.SelectedIndex);
        }
    }
}
