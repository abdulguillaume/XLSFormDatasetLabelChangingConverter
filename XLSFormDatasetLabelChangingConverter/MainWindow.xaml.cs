using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel; 


namespace XLSFormDatasetLabelChangingConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            txtKoboFile.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txtXLSFormFile.Text = txtKoboFile.Text;

            //lblTest.Content = 
            //DateTime.Now.ToFileTime();

            
            //lblTest.Content = converted_file;
        }

        private void btnResultFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel file (*.xls;*.xlsx)|*.xls;*.xlsx| CSV file (*.csv)|*.csv";

            //openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == true)
                txtKoboFile.Text = openFileDialog.FileName;
            //txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
        }

        private void btnXLFFormFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel file (*.xls;*.xlsx)|*.xls;*.xlsx";

            if (openFileDialog.ShowDialog() == true)
                txtXLSFormFile.Text = openFileDialog.FileName;

        }

        private void btnConvertion_Click(object sender, RoutedEventArgs e)
        {
            const int DATA_LABEL_POSITION = 3;
            const int DATA_TYPE_POSITION = 1;

            //ProgressIndicator.IsBusy = true;
            string multiple_choice_str = null;
            /*
                Task.Factory.StartNew(() =>
                {*/

                    //- start of task

                    //https://coderwall.com/p/app3ya/read-excel-file-in-c
                    //Create COM Objects. Create a COM object for everything that is referenced
                    Excel.Application xlApp = new Excel.Application();

                    //kobo workbook
                    Excel.Workbook xlWorkbook_Results = xlApp.Workbooks.Open(txtKoboFile.Text);
                    Excel._Worksheet xlWorksheet_Dataset = xlWorkbook_Results.Sheets[1];
                    Excel.Range xlRange_Dataset = xlWorksheet_Dataset.UsedRange;

                    int rowCount_Dataset = xlRange_Dataset.Rows.Count;
                    int colCount_Dataset = xlRange_Dataset.Columns.Count;

                    //Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    var converted_file = string.Format("{0}\\Converted__{1}.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToFileTime());
                    //var xlWorkbook_Converted = xlApp.Workbooks.Open(converted_file);

                    //xlsform workbook
                    Excel.Workbook xlWorkbook_XLSForm = xlApp.Workbooks.Open(txtXLSFormFile.Text);
                    //survey sheet
                    Excel._Worksheet xlWorksheet_Survey = xlWorkbook_XLSForm.Sheets[1];
                    Excel.Range xlRange_Survey = xlWorksheet_Survey.get_Range("B:B", Type.Missing);

                    //choices sheet
                    Excel._Worksheet xlWorksheet_Choices = xlWorkbook_XLSForm.Sheets[2];
                    Excel.Range xlRange_Choices = xlWorksheet_Choices.get_Range("A:B", Type.Missing);

                    // int rowCount_Survey = xlRange_Survey.Rows.Count;
                    // int rowCount_Choices = xlRange_Choices.Rows.Count;


                    // int tmp = findCodeRowIndex(xlRange_Survey, "cur_ward");


                    //start decoding labels
                    //through the dataset columns
                    int i = 1;
                    for (var j = 1; j < colCount_Dataset; j++)
                    {
                        string toReplace;
                        //first row has the header => Cells(1,j)
                        string header_name =  xlRange_Dataset.Cells[i, j].Value.ToString();

                        string[] h = header_name.Split('/');

                        //toReplace = h.Length == 1? h[0]: h[1];

                        toReplace = h[h.Length - 1];

                        //xlRange_Choices.AutoFilter(1, "");

                        int rowIndex;

                        if (h.Length == 3 && !string.IsNullOrEmpty(multiple_choice_str))//type.StartsWith("select_multiple"))
                        {
                            //search string in choices workbook if multiple choice
                            xlRange_Choices.AutoFilter(1, multiple_choice_str);

                            rowIndex = findCodeRowIndex(
                                xlRange_Choices.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing),
                                                                        toReplace);

                        }
                        else
                        {
                            //search string in survey workbook if not multiple choice 
                            rowIndex = findCodeRowIndex(xlRange_Survey, toReplace);
                            if (!string.IsNullOrEmpty(multiple_choice_str))
                            {
                                multiple_choice_str = null;
                            }
                        }



                        if (rowIndex != -1)
                        {
                            Excel._Worksheet xlWorksheet;

                            if (string.IsNullOrEmpty(multiple_choice_str))
                                xlWorksheet = xlWorksheet_Survey;
                            else
                                xlWorksheet = xlWorksheet_Choices;

                            string label =
                            xlWorksheet.Cells[rowIndex, DATA_LABEL_POSITION].Value.ToString();

                            xlRange_Dataset.Cells[i, j].Value = label;

                            if (xlWorksheet == xlWorksheet_Survey)
                            {
                                string type = xlWorksheet.Cells[rowIndex, DATA_TYPE_POSITION].Value.ToString();

                                if (type.StartsWith("select_multiple"))
                                {
                                    multiple_choice_str = type.Split(' ')[1];
                                }
                                else if (type.StartsWith("select_one"))
                                { 
                                    string single_choice_str = type.Split(' ')[1];
                                    //changeLabelOfRowsBelow()
                                    //xlWorksheet_Choices.AutoFilterMode = false; 
                                    bool tmp = xlRange_Choices.AutoFilter(1, single_choice_str);
                         
                                    var a = 1;
                                    do
                                    {
                                        a++;
                                        try
                                        {
                                            string cell = xlWorksheet_Dataset.Cells[a, j].Value.ToString();
                                            int row = string.IsNullOrEmpty(cell)?-1:findCodeRowIndex(
                                                    xlRange_Choices.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing),
                                                                                                                    cell);
                                            if (row != -1)
                                            {
                                                string lbl = xlWorksheet_Choices.Cells[row, DATA_LABEL_POSITION].Value.ToString();
                                                xlRange_Dataset.Cells[a, j].Value = lbl;
                                            }

                                        }
                                        catch (Exception ex){
                                           // if (a == 2 && j == 1)
                                             //   MessageBox.Show(ex.Message);
                                        }
                                        
                                        Dispatcher.Invoke(DispatcherPriority.Normal, new Action(()=>{
                                       // ProgressIndicator.BusyContent = string.Format(" Column #{0}", j.ToString());
                                            lblTest.Content = string.Format(" Column {0} / Row {1}", j, a);

                                        }));


                                    } while (a < rowCount_Dataset);

                                    
                                }



                            }

                            
                            


                        }

                    }

                    xlWorkbook_Results.SaveAs(converted_file);

                    MessageBox.Show("File Saved!");
                    //int rowCount = xlRange_Dataset.Rows.Count;

                    //int rowCount2 = xlRange_Choices.Rows.Count;

                    //lblTest.Content = string.Format("row count: {0} / {1}", rowCount, rowCount2);





                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //rule of thumb for releasing com objects:
                    //  never use two dots, all COM objects must be referenced and released individually
                    //  ex: [somthing].[something].[something] is bad

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(xlRange_Dataset);
                    Marshal.ReleaseComObject(xlWorksheet_Dataset);

                    Marshal.ReleaseComObject(xlRange_Survey);
                    Marshal.ReleaseComObject(xlWorksheet_Survey);

                    Marshal.ReleaseComObject(xlRange_Choices);
                    Marshal.ReleaseComObject(xlWorksheet_Choices);

                    //close and release
                    xlWorkbook_Results.Close();
                    Marshal.ReleaseComObject(xlWorkbook_Results);

                    xlWorkbook_XLSForm.Close();
                    Marshal.ReleaseComObject(xlWorkbook_XLSForm);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);


            /*
                    //- end of task
                }).ContinueWith((task) => {
                    ProgressIndicator.IsBusy = false;
                }, TaskScheduler.FromCurrentSynchronizationContext());

            */
                

                



        }

        private static int findCodeRowIndex(Excel.Range xlRange, string search)
        {
            int code = -1;
            try
            {
                var currentFind = xlRange.Find(search,
                 xlRange.Cells[1, 1],
                 Excel.XlFindLookIn.xlValues,
                 Excel.XlLookAt.xlPart,
                 Excel.XlSearchOrder.xlByRows,
                 Excel.XlSearchDirection.xlNext,
                 false,
                 false,
                 false);

                code = currentFind.Row;

            }
            catch { 
                
            }

            return code;
            

           // string sAddress = currentFind.Row.ToString();//.Address;



            //~~> Display the found Address
            //MessageBox.Show(sAddress).ToString();
        }
    }
}
