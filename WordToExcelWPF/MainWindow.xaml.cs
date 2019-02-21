using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.IO;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordToExcelWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<string> heading4List = new List<string>();
        List<string> errorList = new List<string>();
        string[] numbersArr = new string[9999];
        int totHours;
        string source;
        string destinationFolder;
        int arrayCounter = 1;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnSource_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.ShowDialog();

            source = openFileDialog.FileName;
            txbInput.Text = source;


        }

        private void BtnDestinationFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog openFileDialog = new FolderBrowserDialog();
            openFileDialog.ShowDialog();
            destinationFolder = openFileDialog.SelectedPath;
        }



        private void ReadWord()
        {
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(txbInput.Text);
            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in document.Paragraphs)
            {
                Microsoft.Office.Interop.Word.Style style = paragraph.get_Style() as Microsoft.Office.Interop.Word.Style;
                string styleName = style.NameLocal;
                string text = paragraph.Range.Text;

                if (styleName == "Heading 2")
                {
                    //Console.WriteLine(text.ToString());
                    AddNumberToList(text);
                    heading4List.Add(text.ToString());

                    arrayCounter++;

                }
                else if (styleName == "Heading 4")
                {
                    //Console.WriteLine(text.ToString());
                    //if (text.Contains("("))
                    //{
                    //    int start = text.IndexOf('(');
                    //    int end = text.IndexOf(')');
                    //    int length = end - start;
                    //    string hour = text.Substring(start+1, length-1);
                    //    int hourInt = int.Parse(hour);
                    //    totHours += hourInt;
                    //}
                    AddNumberToList(text);


                    heading4List.Add("           " + text.ToString());
                    arrayCounter++;

                }
                else if(styleName == "Heading 1")
                {
                    AddNumberToList(text, true, "heading 1");
                }
                else if(styleName == "Heading 3")
                {
                    AddNumberToList(text, true, "heading 3");
                }
                else
                {
                    AddNumberToList(text, true, "everywhere else");
                }

            }

            document.Close();
        }

        private void AddNumberToList(string text, bool comingFromWrong = false, string fromWhere = "")
        {
            if (text.Contains('('))
            {
                string after = text.Substring(text.IndexOf('('));
                string resultString = Regex.Match(after, @"\d+").Value;
                if (resultString != "")
                {
                    if (comingFromWrong)
                    {
                        errorList.Add(fromWhere + " " +text);
                    }
                    else
                    {
                        numbersArr[arrayCounter] = resultString;
                        int hour = int.Parse(resultString);
                        totHours += hour;
                    }
     
                }
            }
        }

        private void CreateExcel()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");

                var headerRow = new List<string[]>()
                {
                    new string[] {"Epic"}
                };

                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                int counter = 1;
                foreach (string heading4 in heading4List)
                {
                    worksheet.Cells["A" + counter].Value = heading4;
                    counter++;
                }

                counter = 1;
                for (int i = 0; i < 999; i++)
                {
                    if (numbersArr[i] != null)
                    {
                        worksheet.Cells["B" + i].Value = numbersArr[i];
                    }
                }

                int counterError = 1;
                foreach (string error in errorList)
                {
                    worksheet.Cells["C" + counterError].Value = error;
                    counterError++;
                }

                //worksheet.Cells["A" + counter].Value = totHours.ToString();

                FileInfo excelFile = new FileInfo(destinationFolder + @"\" + txbOutput.Text);
                excel.SaveAs(excelFile);

            }

        }

        private async void BtnConvert_Click_1(object sender, RoutedEventArgs e)
        {
            heading4List = new List<string>();
            numbersArr = new string[9999];
            arrayCounter = 1;
            totHours = 0;
            errorList = new List<string>();

            lblWorking.Visibility = Visibility.Visible;


            ReadWord();

            CreateExcel();
            lblWorking.Visibility = Visibility.Hidden;

            System.Windows.MessageBox.Show("Done");

        }
    }
}