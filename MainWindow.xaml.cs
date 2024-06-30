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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
//using System.Web;

//using System.Web.UI;


namespace BirdsProject1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        //public MainWindow()
        //{
        //    InitializeComponent();
        //}

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
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


        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string UserName = txtUserName.Text;
            string Password = txtPassword.Password;
            string exelName = "ProjectUsers.xlsx";
            string fileName = Directory.GetCurrentDirectory() + "\\" + exelName;
            //string fileName = "C:\\Users\\97258\\Desktop\\BirdsProject1 (5) (1)\\BirdsProject1\\ProjectUsers.xlsx";
            bool found = SearchExcel(fileName, UserName, Password);
            if (found)
            {
                AfterLogin afterLoginWindow = new AfterLogin();
                afterLoginWindow.Show();
                this.Close();
            }
        }

        public static Boolean SearchExcel(string filePath, string UserName, string Password)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                string nameInCell = usedRange.Cells[row, 1].Value?.ToString(); // Assuming name is in column 1
                string passwordInCell = usedRange.Cells[row, 2].Value?.ToString(); // Assuming password is in column 2

                if (nameInCell == UserName)
                {
                    if (passwordInCell == Password)
                    {
                        workbook.Close();
                        excel.Quit();
                        ReleaseObject(worksheet);
                        ReleaseObject(workbook);
                        ReleaseObject(excel);
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("The password is incorrect, please try again");
                        return false;
                    }
                }
            }

            MessageBox.Show("The user name does not found, please try again");
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

        private void btnSignUpMainFrame_Click(object sender, RoutedEventArgs e)
        {
            SignUp sign = new SignUp();
            sign.Show();
            this.Close();

        }
    }
}
