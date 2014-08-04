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
    /// Company.xaml 的交互逻辑
    /// </summary>
    public partial class Company : MetroWindow
    {
        public Company()
        {
            InitializeComponent();
        }

        DatabaseContext db = new DatabaseContext();
        ObservableCollection<Base.Company> data;

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ProgressRing_company.IsActive = true;
            Task.Factory.StartNew(() =>
            {
                db.SaveChanges();
            });
        }

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                db.Companys.Load();
                data = db.Companys.Local;
                this.Dispatcher.Invoke(new Action(() =>
                {
                    DataGrid_company.ItemsSource = data;
                    DataGrid_company.Visibility = Visibility.Visible;
                    ProgressRing_company.IsActive = false;
                }));
            });
        }

        private void MenuItem_delete_Click(object sender, RoutedEventArgs e)
        {
            data.RemoveAt(DataGrid_company.SelectedIndex);
        }
    }
}
