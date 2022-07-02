using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace CoverterApplication
{
    /// <summary>
    /// Interaction logic for Comparing.xaml
    /// </summary>
    public partial class Comparing : Window
    {
        public Comparing()
        {
            InitializeComponent();
            getColumns();
            logger.Text = "Ընտրեք 2 սյուն և սեղմեք համեմատել";

        }

        private void getColumns()
        {
            ((MainWindow)System.Windows.Application.Current.MainWindow).table1Columns.ForEach(column =>
            {
                file1Columns.Items.Add(new TextBlock { Text = column ,});
            });

            ((MainWindow)System.Windows.Application.Current.MainWindow).table2Columns.ForEach(column =>
            {
                file2Columns.Items.Add(new TextBlock { Text = column });
            });

        }
        private void Button_Cancel_Click(object sender, RoutedEventArgs e)
        {
            
            ((MainWindow)System.Windows.Application.Current.MainWindow).Show();
            Close();
        }

        private void Button_Compare_Click(object sender, RoutedEventArgs e)
        {
            if (file1Columns.SelectedItem is null || file2Columns.SelectedItem is null)
            {
                System.Windows.Forms.MessageBox.Show($"Ընտրեք 2 սյուն ", "ERROR", System.Windows.Forms.MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                disableButtons();
                comparing();
                ((MainWindow)System.Windows.Application.Current.MainWindow).Show();
                Close();

            }
        }

        private void disableButtons() {
            cancelButton.IsEnabled = false;
            compareButton.IsEnabled = false;
            logger.Text = "Սպասեք․․․";
        
        }

        private void comparing()
        {
            string table1OnlyQuery = 
                "SELECT * FROM " +
                $"[converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table1Name}] t1 " +
                $"LEFT JOIN [converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table2Name}] " +
                $"t2 ON t2.{((TextBlock)file2Columns.SelectedItem).Text} = t1.{((TextBlock)file2Columns.SelectedItem).Text}" +
                $" WHERE t2.{((TextBlock)file2Columns.SelectedItem).Text} IS NULL";


            string table2OnlyQuery = 
                $"SELECT * FROM " +
                $"[converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table2Name}]" +
                $" t2 LEFT JOIN [converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table1Name}]" +
                $" t1 ON t1.[{((TextBlock)file1Columns.SelectedItem).Text}] = t2.[{((TextBlock)file2Columns.SelectedItem).Text}]" +
                $" WHERE t1.[{((TextBlock)file1Columns.SelectedItem).Text}] IS NULL";


            string innerQuery = 
                "SELECT * FROM " +
                $"[converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table1Name}] t1 " +
                $"LEFT JOIN [converterAPP].[dbo].[{ ((MainWindow)System.Windows.Application.Current.MainWindow).table2Name}] " +
                $"t2 ON t2.{((TextBlock)file2Columns.SelectedItem).Text} = t1.{((TextBlock)file2Columns.SelectedItem).Text}" +
                $" WHERE t2.{((TextBlock)file2Columns.SelectedItem).Text} IS NOT NULL";

            Console.WriteLine("T1 Only");
            ((MainWindow)System.Windows.Application.Current.MainWindow).sql(table1OnlyQuery,0,true);
            Console.WriteLine("T2 Only");
            ((MainWindow)System.Windows.Application.Current.MainWindow).sql(table2OnlyQuery,0,true);
            Console.WriteLine("Inner");
            ((MainWindow)System.Windows.Application.Current.MainWindow).sql(innerQuery,0,true);
            
        }
    }
}
