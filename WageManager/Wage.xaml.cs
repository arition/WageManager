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
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace WageManager
{
    /// <summary>
    /// Wage.xaml 的交互逻辑
    /// </summary>
    public partial class Wage : MetroWindow
    {
        private class wageRoundConverter : IValueConverter
        {
            public object Convert(Object value, Type targetType, Object parameter, System.Globalization.CultureInfo culture)
            {
                DateTime dt = (DateTime)value;
                return dt.Year + "-" + dt.Month;
            }
            public Object ConvertBack(Object value, Type targetType, Object parameter, System.Globalization.CultureInfo culture)
            {
                string[] s = ((string)value).Split('-');
                return new DateTime(int.Parse(s[0]), int.Parse(s[1]), 1);
            }
        }

        public Wage()
        {
            InitializeComponent();
        }

        DatabaseContext db = new DatabaseContext();
        ObservableCollection<Base.Wage> WageList = new ObservableCollection<Base.Wage>();

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                ObservableCollection<Base.Company> companys = db.Companys.Local;
                db.Employees.ToList().ForEach(d =>
                {
                    Base.Wage wage_old = d.Wages.Where(o => (o.wageRound.Year == DateTime.Now.Year) && (DateTime.Now.Month - o.wageRound.Month == 2)).FirstOrDefault();
                    Base.Wage wage;
                    if (wage_old == null)
                    {
                        wage = new Base.Wage() { employee = d, baseSalary = d.基础工资 };
                    }
                    else
                    {
                        wage = new Base.Wage()
                        {
                            employee = d,
                            baseSalary = d.基础工资,
                            company = wage_old.company,
                            company_tax = wage_old.company_tax,
                            jobSalary = wage_old.jobSalary,
                            houseBonus = wage_old.houseBonus,
                            mealBonus = wage_old.mealBonus
                        };
                    }
                    wage.wageRound = DateTime.Now.AddMonths(-1);
                    WageList.Add(wage);
                });
                this.Dispatcher.Invoke(new Action(() =>
                {
                    ComboBox_employeeid.ItemsSource = WageList;
                    ComboBox_companyid.ItemsSource = companys;
                    ComboBox_companyid_tax.ItemsSource = companys;
                    ProgressRing_wage.IsActive = false;
                    DockPanel_wage.IsEnabled = true;
                }));
                //var ww = db.Employees;
            });
        }

        private void ComboBox_employeeid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Base.Wage currentWage = ComboBox_employeeid.SelectedItem as Base.Wage;
            TextBox_wageRound.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("wageRound"), Source = currentWage, Converter = new wageRoundConverter() });
            ComboBox_companyid.SetBinding(ComboBox.SelectedValueProperty, new Binding() { Path = new PropertyPath("company"), Source = currentWage });
            ComboBox_companyid_tax.SetBinding(ComboBox.SelectedValueProperty, new Binding() { Path = new PropertyPath("company_tax"), Source = currentWage });
            TextBox_baseSalary.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("baseSalary"), Source = currentWage });
            TextBox_jobSalary.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("jobSalary"), Source = currentWage });
            TextBox_performanceBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("performanceBonus"), Source = currentWage });
            TextBox_projectBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("projectBonus"), Source = currentWage });
            TextBox_saleBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("saleBonus"), Source = currentWage });
            TextBox_attendanceBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("attendanceBonus"), Source = currentWage });
            TextBox_overtimeBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("overtimeBonus"), Source = currentWage });
            TextBox_absenceSalary.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("absenceSalary"), Source = currentWage });
            TextBox_adjustmentSalary.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("adjustmentSalary"), Source = currentWage });
            TextBox_socialWelfareDeduction.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("socialWelfareDeduction"), Source = currentWage });
            TextBox_publicFundDeduction.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("publicFundDeduction"), Source = currentWage });
            TextBox_adjustmentDeduction.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("adjustmentDeduction"), Source = currentWage });
            TextBox_mealBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("mealBonus"), Source = currentWage });
            TextBox_houseBonus.SetBinding(TextBox.TextProperty, new Binding() { Path = new PropertyPath("houseBonus"), Source = currentWage });
        }

        private void btm_preview_Click(object sender, RoutedEventArgs e)
        {
            foreach (Base.Wage wage in WageList)
            {
                db.Wages.Add(wage);
            }
            db.SaveChanges();
        }
    }
}
