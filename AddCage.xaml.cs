using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;

namespace BirdsProject1
{
    /// <summary>
    /// Interaction logic for AddCage.xaml
    /// </summary>
    public partial class AddCage : System.Windows.Window
    {
        public AddCage()
        {
            InitializeComponent();
            Loaded += AddCage_Loaded;
        }

        private void AddCage_Loaded(object sender, RoutedEventArgs e)
        {
            // Disable the Maximize button on the window
            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            var style = NativeMethods.GetWindowLong(hwnd, NativeMethods.GWL_STYLE);
            style &= ~NativeMethods.WS_MAXIMIZEBOX;
            NativeMethods.SetWindowLong(hwnd, NativeMethods.GWL_STYLE, style);
        }

        internal static class NativeMethods
        {
            public const int GWL_STYLE = -16;
            public const int WS_MAXIMIZEBOX = 0x10000;

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern int GetWindowLong(IntPtr hwnd, int index);

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern int SetWindowLong(IntPtr hwnd, int index, int value);
        }

        private void btnAddCage_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;
            bool isLetter = false;
            bool isDigit = false;

            string exelNameCage = "cages.xlsx";
            string fileCage = Directory.GetCurrentDirectory() + "\\" + exelNameCage;

            foreach (char letter in txtSerialNumberCage.Text)
            {
                if (char.IsLetter(letter))
                    isLetter = true;
                else if (char.IsDigit(letter))
                    isDigit = true;
                if (isLetter && isDigit)
                    break;
            }
            bool goodCage = isLetter && isDigit;
            if (!goodCage)
            {
                MessageBox.Show("The cage number must contains at least one letter, one digit");
                flag = false;
            }
            
            if (SearchExcel(fileCage, txtSerialNumberCage.Text, 1))
            {
                MessageBox.Show("The cage already exists");
                flag = false;
            }

            if (!check_validation(txtLengthCage.Text, "length"))
                flag = false;
            if (!check_validation(txtWidthCage.Text, "width"))
                flag = false;
            if (!check_validation(txtHeightCage.Text,"height"))
                flag = false;
            if(int.Parse(txtHeightCage.Text) > 3000)
            {
                MessageBox.Show("The height must be under 3001 meters");
                flag = false;
            }

            if (cmbMaterial.Text == "" || cmbMaterial.Text == " ")
            {
                MessageBox.Show("You must choose a material");
                flag = false;
            }

            if (flag)
            {
                Cage newCage = new Cage(txtSerialNumberCage.Text, int.Parse(txtLengthCage.Text), int.Parse(txtWidthCage.Text), int.Parse(txtHeightCage.Text),cmbMaterial.Text);
                WriteToExcel(fileCage, txtSerialNumberCage.Text, txtLengthCage.Text, txtWidthCage.Text, txtHeightCage.Text, cmbMaterial.Text);
                MessageBox.Show("The cage was added successfully");
                AfterLogin afterLoginWindow = new AfterLogin();
                afterLoginWindow.Show();
                this.Close();
            }
        }

        public static void WriteToExcel(string filePath, string serialNumberCage, string LengthCage, string WidthCage, string HeightCage, string Material)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            // Find the first empty row to write the data
            int newRow = rowCount + 1;

            // Write the data to the worksheet
            usedRange.Cells[newRow, 1].Value = serialNumberCage; // Serial number
            usedRange.Cells[newRow, 2].Value = LengthCage; // Species
            usedRange.Cells[newRow, 3].Value = WidthCage; // Subspecies
            usedRange.Cells[newRow, 4].Value = HeightCage; // Date
            usedRange.Cells[newRow, 5].Value = Material; // Gender

            workbook.Save();
            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
        }
        public static bool SearchExcel(string filePath, string Value, int c)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                string Cell = usedRange.Cells[row, c].Value?.ToString(); // Assuming serial number is in column 3

                if (Cell == Value)
                {
                    workbook.Close();
                    excel.Quit();
                    ReleaseObject(worksheet);
                    ReleaseObject(workbook);
                    ReleaseObject(excel);
                    return true;
                }
            }

            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
            return false;
        }
        static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception occurred while releasing object: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        public Boolean check_validation(string val,string type)
        {
            bool flag = true;

            if (int.TryParse(val, out int value))
            {
                if (value < 0)
                {
                    MessageBox.Show("The " + type + " must a positive number");
                    flag = false;
                }
            }
            else
            {
                MessageBox.Show("The " + type + " must a valid number");
                flag = false;
            }

            return flag;
        }

        private void btnAfterLogin_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }

        private void btnMainWindow_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}
