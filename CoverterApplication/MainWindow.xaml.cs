using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Threading;
using MessageBox = System.Windows.Forms.MessageBox;

namespace CoverterApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SqlConnection myConn = new SqlConnection("Server=(localdb)\\mssqllocaldb;Database=converterAPP;Trusted_Connection=True;MultipleActiveResultSets=true;");
        SqlConnection databaseConnection = new SqlConnection("Server=(localdb)\\mssqllocaldb;Trusted_Connection=True;MultipleActiveResultSets=true;");

        public MainWindow()
        {
            InitializeComponent();
            loading("Ընտրեք 2 Excel ֆայլ և սեղմեք սկսել:");
            createDatabase();
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }
        private void createDatabase()
        {
            string databaseCreateCmd = "IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'converterAPP')CREATE DATABASE converterAPP";

            SqlCommand cmd = new SqlCommand(databaseCreateCmd, databaseConnection);
            try
            {
                databaseConnection.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                databaseConnection.Close();
            }
        }

        private void start(object sender, RoutedEventArgs e)
            
        {


            if (String.IsNullOrEmpty(textBox1.Text) || String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Ընտրեք 2 ֆայլ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                disableButtons();
                writeExcel(textBox1.Text);
                writeExcel(textBox2.Text);
                enableButtons();
            }


        }
        private void disableButtons()
        {
            logger.Background = Brushes.Gray;            
            startBtn.IsEnabled = false;
            file1Btn.IsEnabled = false;
            file2Btn.IsEnabled = false;
            startBtn.Content = "Սպասեք․․․";
            loading("Սկսվից․․․");
        }

        private void enableButtons()
        {
            changeBusy();            
            logger.Background = Brushes.Green;
            logger.Foreground = Brushes.White;
            file1Btn.IsEnabled = true;
            file2Btn.IsEnabled = true;
            loading("Տվյալները հաջողությամբ ավելացվեցին");
        }
        public void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame(true);
            Dispatcher.CurrentDispatcher.BeginInvoke
            (
            DispatcherPriority.Background,
            (SendOrPostCallback)delegate (object arg)
            {
                var f = arg as DispatcherFrame;
                f.Continue = false;
            },
            frame
            );
            Dispatcher.PushFrame(frame);
        }
         void loading(string message)
        {
             logger.Text = message;
            DoEvents();
            

        }
        public void sql(string command)
        {
            SqlCommand myCommand = new SqlCommand(command, myConn);
            try
            {
                myConn.Open();
                myCommand.ExecuteNonQuery();
                //MessageBox.Show("DataBase is Created Successfully", "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error");
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }
        }

        private void changeBusy()
        {
            this.Dispatcher.Invoke(() =>
            {
                startBtn.IsEnabled = true;
            });

        }
        private void file_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                if (((System.Windows.Controls.Button)sender).Name == "file1Btn")
                {
                    textBox1.Text = filename;
                }
                else
                {
                    textBox2.Text = filename;
                }


            }
        }



        void writeExcel(string fileName, bool firstRowIsHeader = true)
        {
            try
            {

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
                {
                    string fileShortName = fileName.Substring(fileName.LastIndexOf("\\")+1);
                    string tableName = $"Table_{fileShortName.Substring(0,fileShortName.IndexOf("."))}_{DateTime.UtcNow.Second}_{DateTime.UtcNow.Millisecond}";
                    string createCmd = $"if not exists" +
                        $" (select * from sysobjects where name='{tableName}' and xtype='U')" +
                        $"create table {tableName}" +
                        $"(ID INT NOT NULL IDENTITY(1,1) PRIMARY KEY";

                    //Read the first Sheets 
                    Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    Row[] rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>().ToArray();
                    int counter = 0;
                    for(int a = 0;a<rows.Length;a++)
                    {
                        string insertCmd = $"Insert Into [converterAPP].[dbo].[{tableName}] VALUES(";
                        counter = counter + 1;

                        if (counter == 1)
                        {
                            var j = 1;
                            foreach (Cell cell in rows[a].Descendants<Cell>())
                            {
                                var columnName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                                columnName = normalizedColumnName(columnName,j);
                                createCmd += $",{columnName} nvarchar(max)";
                            }
                            createCmd += ");";
                            Console.WriteLine(createCmd);
                            sql(createCmd);
                            loading("Աղյուսակը ստեղծվեց");
                        }
                        else
                        {
                            int i = 0;
                            foreach (Cell cell in rows[a].Descendants<Cell>())
                            {
                                string cellValue = GetCellValue(doc, cell);
                                string nullString = "-";
                                string cmdData = String.IsNullOrEmpty(cellValue) ? nullString : cellValue;
                                
                                insertCmd += $"N'{cmdData}',";
                                Console.Write($" {GetCellValue(doc, cell)} ");
                                i++;
                            }
                            Console.WriteLine();
                            insertCmd = insertCmd.Remove(insertCmd.Length - 1);
                            insertCmd += ")";

                            sql(insertCmd);
                            loading($"Տվյալի ավելացում - {a} | Ֆայլ - {fileShortName} ");
                        }
                    }
                    

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }//doc write end

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell?.CellValue?.InnerText;
            
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var result = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                return result;
            }
            return value;
        }

        private string normalizedColumnName(string name,int j)
        {
            string normalizedName = name;
            if (String.IsNullOrWhiteSpace(name))
            {
                return $"Field{j}";
            }
            if (int.TryParse($"{normalizedName[0]}",out int value))
            {
                normalizedName = $"Field_{normalizedName}";
            }

            string[] chars = new string[] {"\n"," ",  ",", ".", "/","\\", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";","-", "_", "(", ")", ":", "|", "[", "]" };
            
            for (int i = 0; i < chars.Length; i++)
            {
                if (normalizedName.Contains(chars[i]))
                {
                    normalizedName = normalizedName.Replace(chars[i], "");
                }
            }
            
            return normalizedName;
        }

    }
}
