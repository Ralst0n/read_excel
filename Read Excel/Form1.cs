using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace Read_Excel
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;

        string timeDirectory;
        string expDirectory;
        string timePdfsPath;
        string expPdfsPath;
        string myWorkbook;

        public Form1()
        {
            InitializeComponent();
            string savePath;
            string[] args = Environment.GetCommandLineArgs();
            //if sufficient arguments aren't provided, let user know and quit program.
            if (args.Length < 4)
            {
                MessageBox.Show($"4 arguments required {args.Length} provided.");
                Application.Exit();
                Environment.Exit(1);
            }
            //add the parts of the path that got separated back together if args is greater then 4. The last two arguments are known directories so they should never need to be rejoined. 
            if (args.Length > 4)
            {
                for (int i = 2; i < (args.Length-2); i++)
                {
                    args[1] += " " + args[i];
                }
            }
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath($"{args[1]}"));
            xlWorksheet = xlWorkbook.Worksheets["backup_list"];
            savePath = Path.GetDirectoryName(args[1]);
            timePdfsPath = args[args.Length - 2] + @"\";
            expPdfsPath = args[args.Length - 1] + @"\";
            xlRange = xlWorksheet.UsedRange;
            Directory.CreateDirectory(savePath + @"\tsbackup");
            Directory.CreateDirectory(savePath + @"\expbackup");
            timeDirectory = savePath + @"\tsbackup\";
            expDirectory = savePath + @"\expbackup\";
            myWorkbook = args[1];
            Shown += Form1_Shown;

        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            //Declare an array to hold the two words that matter, then set both sections equal to false
            
            label1.Text = "Loading...";
            string[] keyWords = new string[2] { "Time Sheets", "Expense Reports" };
            bool timeSheet = false;
            bool expenseSheet = false;
            progressBar1.Maximum = xlRange.Rows.Count -1;
            for (int i = 1; i < xlRange.Rows.Count; i++)
            {


                //expense sheet first since that word comes  second
                //sets expenseSheet true and timeSheet false if we get into expense section
                //when the current cell says expense sheet
                if (Convert.ToString(xlRange.Cells[i, 1].Value2) == keyWords[1])
                {
                    expenseSheet = true;
                    timeSheet = false;
                }
                //sets timesheet true when we get into the timesheet section
                //when the current cell says timehseet
                else if (Convert.ToString(xlRange.Cells[i, 1].Value2) == keyWords[0])
                {
                    timeSheet = true;

                }
                //sets expense to false after the last expense report person is read
                else if (expenseSheet & String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 1].Value2)))
                {
                    expenseSheet = false;
                }
                    

                //if the cell isn't blank and timesheet is false, 
                else if (!String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 1].Value2)))
                {
                    if ((timeSheet || expenseSheet) & Convert.ToString(xlRange.Cells[i, 3].Value2) != "Employee Name")
                    {
                        //take individual employee name string and the period listed along with name
                        //conversion to properly formatted string credit to Yuck of stack overflow
                        String EmployeeName = Convert.ToString(xlRange.Cells[i, 3].Value2);
                        DateTime periodEndingDate = DateTime.FromOADate(xlRange.Cells[i, 1].Value2);
                        //MessageBox.Show($"{periodEndingDate}");
                        String periodStart = periodEndingDate.AddDays(-13).ToString("yy-MM-dd");
                        String periodEnd = periodEndingDate.ToString("yy-MM-dd");
                        String pdfDates = $"{periodStart} - {periodEnd}.pdf";
                        if (timeSheet)
                        {
                            //Console.WriteLine($"{i}: {Convert.ToString(xlRange.Cells[i, 3].Value2)} is TimeSheet {pdfDates}");
                            PdfManipulation pdf = new PdfManipulation(timePdfsPath + pdfDates);
                            //run pdfmanipulations find page method on employee last name.
                            label1.Text = $"Searching for {EmployeeName} in {pdfDates}...";
                            int[] pages = pdf.FindPages(EmployeeName.Split(null)[1], timeSheet);
                            if (pages[0] == 0 && pages[1] == 0)
                            {
                                xlRange.Cells[i, 6].Value2 = $"Could not find {EmployeeName} in {pdfDates}.pdf";
                            }
                            else
                            {
                                pdf.AddToPdf(pages, timeDirectory + $"{EmployeeName} {periodEnd}.pdf");
                                xlRange.Cells[i, 6].Value2 = $"Found on pages {pages[0]} - {pages[1]} of {pdfDates}";
                            }
                        }
                        else if (expenseSheet)
                        {
                            PdfManipulation pdf = new PdfManipulation(expPdfsPath + pdfDates);
                            label1.Text = $"Searching for {EmployeeName} in {pdfDates}...";
                            int[] pages = pdf.FindPages(EmployeeName.Split(null)[1], timeSheet);
                            if (pages[0] == 0 && pages[1] == 0)
                            {
                                xlRange.Cells[i, 6].Value2 = $"Could not find {EmployeeName} in {pdfDates}.pdf";
                            }
                            else
                            {
                                pdf.AddToPdf(pages, expDirectory + $"{EmployeeName} {periodEnd}.pdf");
                                xlRange.Cells[i, 6].Value2 = $"Found on pages {pages[0]} - {pages[1]} of {pdfDates}";
                                //Console.WriteLine($"{ i}: {Convert.ToString(xlRange.Cells[i, 3].Value2)} is ExpenseSheet {pdfDates}");
                            }
                        }
                    }
                }
                progressBar1.Value += 1;
                
            }
            //MessageBox.Show($"progress bar got to {progressBar1.Value} out of {progressBar1.Maximum} as Row Count was {xlRange.Rows.Count}");
            label1.Text = $"Attempting to open your workbook";
            ReleaseWorkbook();
            System.Diagnostics.Process.Start(myWorkbook);
            Application.Exit();
            Environment.Exit(0);

        }

        public void ReleaseWorkbook()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Save();
            xlWorkbook.Close(true, null, null);
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Invoke(new Action(() => { MessageBox.Show("PDF Finder v3.0\nAuthor: Ralston Lawson\n\u00A9 2017", "About"); }));
        }
    }
}
