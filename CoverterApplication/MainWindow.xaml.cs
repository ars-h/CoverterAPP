using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using IronXL;
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
        public SqlConnection myConn = new SqlConnection("Server=(localdb)\\mssqllocaldb;Database=converterAPP;Trusted_Connection=True;MultipleActiveResultSets=true;");
        public SqlConnection databaseConnection = new SqlConnection("Server=(localdb)\\mssqllocaldb;Trusted_Connection=True;MultipleActiveResultSets=true;");
        public List<string> table1Columns;
        public List<string> table2Columns;
        public string table1Name;
        public string table2Name;
        public MainWindow()
        {
            InitializeComponent();
            loading("Ընտրեք 2 Excel ֆայլ և սեղմեք սկսել");
            createDatabase();
            table1Columns = new List<string> { };
            table2Columns = new List<string> { };
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
                MessageBox.Show($"Ընտրեք 2 ֆայլ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                disableButtons();
                writeExcel(textBox1.Text, 1);
                writeExcel(textBox2.Text, 2);
                enableButtons();
                if ((bool)doComparing.IsChecked)
                {
                    Comparing comparingWindow = new Comparing();
                    comparingWindow.Show();
                    Hide();
                }
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
            startBtn.Content = "Սկսել";
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
        public void sql(string command, int tableNumber = 0, bool comparing = false)
        {
            SqlCommand myCommand = new SqlCommand(command, myConn);

            try
            {
                if (comparing)
                {
                    myConn.Open();
                    SqlDataReader reader = myCommand.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read()) // построчно считываем данные
                        {


                            for (int i = 1; tableNumber == 1 ? i <= table1Columns.Count() : i <= table2Columns.Count(); i++)
                            {
                                Console.Write(reader.GetValue(i));
                                Console.Write("   |   ");
                            }

                            Console.WriteLine();
                        }

                    }
                }
                else
                {
                    myConn.Open();
                    myCommand.ExecuteNonQuery();
                }
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

        private void compare()
        {

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
                    file1Btn.Content = "Ընտրել նոր ֆայլ";
                }
                else
                {
                    textBox2.Text = filename;
                    file2Btn.Content = "Ընտրել նոր ֆայլ";
                }


            }
        }



        private void writeExcel(string fileName, int tableNumber)
        {
            try
            {


                WorkBook workbook = WorkBook.Load(fileName);
                WorkSheet sheet = workbook.DefaultWorkSheet;
                System.Data.DataTable dataTable = sheet.ToDataTable(true);
                string fileShortName = fileName.Substring(fileName.LastIndexOf("\\") + 1);
                fileShortName = fileShortName.Replace(" ", "");
                string tableName = $"Table_{fileShortName.Substring(0, fileShortName.IndexOf("."))}_{DateTime.UtcNow.Second}_{DateTime.UtcNow.Millisecond}";
                if (tableNumber == 1)
                {
                    table1Name = tableName;
                }
                else
                {
                    table2Name = tableName;
                }
                string createCmd = 
                    $"if not exists" +
                    $" (select * from sysobjects where name='{tableName}' and xtype='U')" +
                    $"create table [{tableName}]" +
                    $"(ID INT NOT NULL IDENTITY(1,1) PRIMARY KEY";


                string createDistinctCmd = 
                    $"if not exists" +
                    $" (select * from sysobjects where name='distinct_{tableName}' and xtype='U')" +
                    $"create table [distinct_{tableName}]" +
                    $"(ID INT NOT NULL IDENTITY(1,1) PRIMARY KEY";

                string innerCmd =
                    $"if not exists" +
                    $" (select * from sysobjects where name='inner_{tableName}' and xtype='U')" +
                    $"create table [inner_{tableName}]" +
                    $"(ID INT NOT NULL IDENTITY(1,1) PRIMARY KEY";





                //get columns names and create table
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {

                    string columnName = dataTable.Columns[i].ColumnName;
                    columnName = normalizedColumnName(columnName, i);
                    createCmd += $",[{columnName}] nvarchar(max)";
                    createDistinctCmd += $",[{columnName}] nvarchar(max)";
                    innerCmd += $",[{columnName}] nvarchar(max)";

                    if (tableNumber == 1)
                    {
                        table1Columns.Add(columnName);
                    }
                    else
                    {
                        table2Columns.Add(columnName);

                    }
                    Console.WriteLine(columnName);
                }
                createCmd += ");";
                sql(createCmd);
                
                if ((bool)doComparing.IsChecked)
                {
                    sql(createDistinctCmd);
                    if (tableNumber == 1)
                    {
                        sql(innerCmd);
                    }
                }
                    loading("Աղյուսակը ստեղծվեց");
                //end craeting tables


                //start adding rows
                int rowIndex = 1;
                foreach (DataRow row in dataTable.Rows)
                {
                    string insertCmd = $"Insert Into [converterAPP].[dbo].[{tableName}] VALUES(";
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        var test = row[i].GetType();
                        string cellValue = row[i].ToString();
                        cellValue = string.IsNullOrWhiteSpace(cellValue) ? "null" : cellValue;
                        insertCmd += $"N'{cellValue}',";

                    }
                    insertCmd = insertCmd.Remove(insertCmd.Length - 1);
                    insertCmd += ")";

                    sql(insertCmd);
                    loading($"Տվյալի ավելացում - {rowIndex} | Ֆայլ - {fileShortName} ");
                    rowIndex++;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private string normalizedColumnName(string name, int j)
        {
            string normalizedName = name;
            if (String.IsNullOrWhiteSpace(name))
            {
                return $"Field{j}";
            }
            if (int.TryParse($"{normalizedName[0]}", out int value))
            {
                normalizedName = $"Field_{normalizedName}";
            }

            string[] chars = new string[] { "\n", " ", ",", ".", "/", "\\", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";", "-", "_", "(", ")", ":", "|", "[", "]" };

            for (int i = 0; i < chars.Length; i++)
            {
                if (normalizedName.Contains(chars[i]))
                {
                    normalizedName = normalizedName.Replace(chars[i], "");
                }
            }

            return normalizedName;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }



        /*  void writeExcel(string fileName, bool firstRowIsHeader = true)
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

          }//doc write end*/

        /*private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell?.CellValue?.InnerText;
            
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var result = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                return result;
            }
            return value;
        }*/



    }
}
