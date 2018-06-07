using System;
using System.Windows.Forms;
using FindCell.Services;
using Excel = Microsoft.Office.Interop.Excel; 

namespace FindCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FindWordInExcel(word: "red", fileName: "FindME_before.xlsx");
           
        }

        private void FindWordInExcel(string word, string fileName)
        {

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(Filename: System.IO.Path.GetFullPath("ExcelFiles\\" + fileName));
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];//.get_Item(tabNumber)

                Excel.Range usedRange = xlWorkSheet.UsedRange;
                foreach (Excel.Range rng in usedRange.Cells)
                {
                    if (rng.Value == word)
                        rng.Interior.Color = Excel.XlRgbColor.rgbRed;
                }

                xlWorkBook.SaveAs(Filename: System.IO.Path.GetFullPath("ExcelFiles\\FindME_after.xlsx"));
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                label1.Visible = true;
                
            }
            catch (Exception ex)
            {
                LogFile.WriteToLog(@"Error occured during searching a word in excel file.\nError meesage: " + ex.Message + "\nStackTrace path: " + ex.StackTrace);
                MessageBox.Show(@"Error occured during searching a word in excel file.\nError meesage: " + ex.Message);
            }
            finally
            {
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }
        }

        private static void ReleaseObject(object comObject)
        {
            try
            {
                if (comObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
                    comObject = null;
                }
                
            }
            catch (Exception ex)
            {
                comObject = null;
                LogFile.WriteToLog(@"Error occured while releasing object.\nError meesage: " + ex.Message + "\nStackTrace path: " + ex.StackTrace);
                MessageBox.Show("Error occured while releasing object " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
