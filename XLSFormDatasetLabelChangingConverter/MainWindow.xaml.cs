using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        List<KeyValuePair<string, List<CodeLabel>>> listChoices = null;

        List<KeyValuePair<string, string[]>> listSurvey = null;

        public class CodeLabel
        {
            public string code { get; private set; }
            public string label { get; private set; }

            public CodeLabel(string code, string label)
            {
                this.code = code;
                this.label = label;
            }

            //public string getLabel(string code)
            //{
            //    return this.code == code ? this.label : null;
            //}

        }
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

            ProgressIndicator.IsBusy = true;

            //https://coderwall.com/p/app3ya/read-excel-file-in-c
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            //xlsform workbook
            Excel.Workbook xlWorkbook_XLSForm = xlApp.Workbooks.Open(txtXLSFormFile.Text, ReadOnly: true);
            //survey sheet
            Excel._Worksheet xlWorksheet_Survey = xlWorkbook_XLSForm.Sheets[1];
            Excel.Range xlRange_Survey = xlWorksheet_Survey.UsedRange;

            //choices sheet
            Excel._Worksheet xlWorksheet_Choices = xlWorkbook_XLSForm.Sheets[2];
            Excel.Range xlRange_Choices = xlWorksheet_Choices.UsedRange;


            Task.Factory.StartNew(() =>
              {

                  //- start of task

                  Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                  {
                      ProgressIndicator.BusyContent = "Processing XLSForm file...";
                  }));

                  listChoices = getListFromChoicesWorksheet(xlRange_Choices);
                  listSurvey = getListFromSurveyWorksheet(xlRange_Survey);

                  //cleanup
                  GC.Collect();
                  GC.WaitForPendingFinalizers();
                  Marshal.ReleaseComObject(xlRange_Survey);
                  Marshal.ReleaseComObject(xlWorksheet_Survey);
                  Marshal.ReleaseComObject(xlRange_Choices);
                  Marshal.ReleaseComObject(xlWorksheet_Choices);
                  //close and release
                  xlWorkbook_XLSForm.Close();
                  Marshal.ReleaseComObject(xlWorkbook_XLSForm);

                //- end of task
            }).ContinueWith((task) =>
            {
                ProgressIndicator.IsBusy = false;
            }, TaskScheduler.FromCurrentSynchronizationContext()); 

        }

        private void btnConvertion_Click(object sender, RoutedEventArgs e)
        {
            const int DATA_LABEL_POSITION = 3;
            const int DATA_TYPE_POSITION = 1;

            string multiple_choice_str = null;

            //https://coderwall.com/p/app3ya/read-excel-file-in-c
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            //kobo results workbook
            Excel.Workbook xlWorkbook_Results = xlApp.Workbooks.Open(txtKoboFile.Text, ReadOnly: true);
            Excel._Worksheet xlWorksheet_Dataset = xlWorkbook_Results.Sheets[1];
            Excel.Range xlRange_Dataset = xlWorksheet_Dataset.UsedRange;

            int rowCount_Dataset = xlRange_Dataset.Rows.Count;
            int colCount_Dataset = xlRange_Dataset.Columns.Count;

            var converted_file = string.Format("{0}\\Converted_{1}.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToFileTime());

           // Stopwatch stopWatch = new Stopwatch();

           // stopWatch.Start();


          //Task.Factory.StartNew(() =>
          //  {

          //      //- start of task

          //      Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
          //      {
          //          ProgressIndicator.BusyContent = "Processing file...";
          //      }));
 

                //start decoding labels

                int rowHeader = 1; //this will be fix during the whole process

                for (var j = 1; j < colCount_Dataset; j++)
                {
                    string toReplace;
                    //first row has the header => Cells(1,j)
                    string header_name = null;
                    try
                    {
                        header_name = xlRange_Dataset.Cells[rowHeader, j].Value.ToString();
                    }
                    catch {
                        break;
                    }

                    string[] h = header_name.Split('/');

                    toReplace = h[h.Length - 1];

                    string[] survey_col_info = null;

                    try
                    {
                        if (h.Length <= 2)
                        {
                            survey_col_info = listSurvey.FirstOrDefault(x => x.Key == toReplace).Value;
                            xlRange_Dataset.Cells[rowHeader, j].Value = survey_col_info[2];
                        }
                        else if (h.Length == 3)
                        {
                            //in the dataset results => h[1] survey question code (h[2] choices code/ to use later, h[0] type)
                            survey_col_info = listSurvey.FirstOrDefault(x => x.Key == h[1]).Value;

                            //no important if tmp[0] is select_one or select_multiple
                            if (string.IsNullOrEmpty(survey_col_info[1]))
                            {
                                List<CodeLabel> _codeLabels = listChoices.FirstOrDefault(x => x.Key == survey_col_info[1]).Value;
                                CodeLabel _codeLabel = _codeLabels.Find(x => x.code == h[2]);
                                xlRange_Dataset.Cells[rowHeader, j].Value = _codeLabel.label;

                            }

                        }

                    }
                    catch { }


                    int row = 1; //change data for entire column

                    do
                    {
                        row++;
                        try
                        {
                            string cell = xlWorksheet_Dataset.Cells[row, j].Value.ToString();

                            if (survey_col_info[0] == "select_multiple")
                            {
                                break;
                            }

                            if (string.IsNullOrEmpty(cell) || string.IsNullOrEmpty(survey_col_info[0]))
                                continue;

                            List<CodeLabel> _codeLabels = listChoices.FirstOrDefault(x => x.Key == survey_col_info[1]).Value;
                            CodeLabel _codeLabel = _codeLabels.Find(x => x.code == cell);
                            xlRange_Dataset.Cells[row, j].Value = _codeLabel.label;

                        }
                        catch (Exception ex)
                        {

                        }


                    } while (row < rowCount_Dataset);




                }

                xlWorkbook_Results.SaveAs(converted_file);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange_Dataset);
                Marshal.ReleaseComObject(xlWorksheet_Dataset);

                //close and release
                xlWorkbook_Results.Close();
                Marshal.ReleaseComObject(xlWorkbook_Results);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                //- end of task
            //}).ContinueWith((task) =>
            //{
            //    ProgressIndicator.IsBusy = false;
            //    stopWatch.Stop();

            //    MessageBox.Show(string.Format("{0} minute(s)", stopWatch.Elapsed.Minutes));

            //}, TaskScheduler.FromCurrentSynchronizationContext()); 


            //ProgressIndicator.IsBusy = true;

        }

        private List<KeyValuePair<string, List<CodeLabel>>> getListFromChoicesWorksheet(Excel.Range xlRange)
        {
            const int DATA_LABEL_POSITION = 3;
            const int DATA_CODE_POSITION = 2;
            const int DATA_GRP_POSITION = 1;

            int rowCount = xlRange.Rows.Count;

            List<KeyValuePair<string, List<CodeLabel>>> _list = new List<KeyValuePair<string, List<CodeLabel>>>();

            string codeGrpTmp = null;

            List<CodeLabel> codeLabelGrp = new List<CodeLabel>(); //not initialized

            //a=2 => skip the header
            for(int a = 2; a < rowCount; a++)
            {
	            string codeGrp = xlRange.Cells[a, DATA_GRP_POSITION].Value.ToString();
 
	            if(string.IsNullOrEmpty(codeGrpTmp)) //ok
	            {
		            codeGrpTmp = codeGrp;
	            }
	            else if(codeGrpTmp != codeGrp) //it is in a different group
	            {
                    //let store the old group as key value pair group=>list(code,label)
		            _list.Add(new KeyValuePair<string, List<CodeLabel>>(codeGrpTmp, codeLabelGrp));

                    //start a new group
                    codeLabelGrp = new List<CodeLabel>();

                    //set codeGrpTmp to new group info
                    codeGrpTmp = codeGrp;
		
	            }

                codeLabelGrp.Add(
                         new CodeLabel(
                             xlRange.Cells[a, DATA_CODE_POSITION].Value.ToString(),
                             xlRange.Cells[a, DATA_LABEL_POSITION].Value.ToString()
                         )
                     );
	            
		
            }

            return _list;
        }


        private List<KeyValuePair<string, string[]>> getListFromSurveyWorksheet(Excel.Range xlRange)
        {
            const int DATA_LABEL_POSITION = 3;
            const int DATA_CODE_POSITION = 2;
            const int DATA_TYPE_POSITION = 1;

            int rowCount = xlRange.Rows.Count;

            List<KeyValuePair<string, string[]>> _list = new List<KeyValuePair<string, string[]>>();

            string[] _values = new string[3]; //not initialized

            //a=2 => skip the header
            for (int a = 2; a < rowCount; a++)
            {
                string type = null;

                try
                {
                    type = xlRange.Cells[a, DATA_TYPE_POSITION].Value.ToString();
                }
                catch 
                { 
                
                }

                if (string.IsNullOrEmpty(type) || type.StartsWith("begin") || type.StartsWith("end r") || type.StartsWith("end g"))
                    continue;

                string code = xlRange.Cells[a, DATA_CODE_POSITION].Value.ToString();

                if (type.StartsWith("select"))
                { 
                    string[] tmp = new string[3];
                    var split = type.Trim(' ').Split(' ');
                    tmp[0] = split[0];
                    tmp[1] = split[1];
                    tmp[2] = xlRange.Cells[a, DATA_LABEL_POSITION].Value.ToString();

                    _list.Add(new KeyValuePair<string, string[]>(code, tmp));
                    continue;

                }



                _list.Add(new KeyValuePair<string, string[]>(code,
                        new string[3] { 
                            "", "", xlRange.Cells[a, DATA_LABEL_POSITION].Value.ToString()
                        }
                    ));


            }

            return _list;
        }

        private void donothing()
        {
            return;
            //throw new NotImplementedException();
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
