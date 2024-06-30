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
    /// Interaction logic for SignUp.xaml
    /// </summary>
    public partial class SignUp : System.Windows.Window
    {

        public SignUp()
        {
            InitializeComponent();
            Loaded += SignUp_Loaded;

        }

        private void SignUp_Loaded(object sender, RoutedEventArgs e)
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

        private void btnSingUp_Click(object sender, RoutedEventArgs e)
        {
            int counterNum=0;
            bool flag = true;

            //ID
            if (IDSingUp.Text.Length == 9)
            {
                for (int i = 0; i < IDSingUp.Text.Length; i++)
                    if (!int.TryParse(IDSingUp.Text[i].ToString(), out int result))
                    {
                        MessageBox.Show("The ID must contain only numbers.");
                        flag = false;

                    }
            }
            else
            {
                MessageBox.Show("The ID must be 9 characters long.");
                flag = false;
            }

            //User name
            if (string.IsNullOrEmpty(txtUserNameSignUp.Text))
            {
                MessageBox.Show("The user name can't be empty");
                flag = false;
            }

            if (txtUserNameSignUp.Text.Length >=6 && txtUserNameSignUp.Text.Length <= 8)
            {
                for (int i = 0; i < txtUserNameSignUp.Text.Length; i++)
                {

                    if (char.IsDigit(txtUserNameSignUp.Text[i]))
                    {
                        counterNum++;
                    }
                }

                if (counterNum <= 2)
                {
                    string exelNameUser = "ProjectUsers.xlsx";
                    string filePath = Directory.GetCurrentDirectory() + "\\" + exelNameUser;

                    if (SearchExcel(filePath, txtUserNameSignUp.Text))
                    {
                        MessageBox.Show("The user name is already exists");
                        flag = false;
                    }
                }
                else
                {
                    MessageBox.Show("The user name must contain up to 2 digits");
                    flag = false;
                }
            }
            else
            {
                MessageBox.Show("The user name must contain between 6 and 8 characters");
                flag = false;
            }

            //Password
            if(txtPasswordSignUp.Text.Length >= 8 && txtPasswordSignUp.Text.Length <= 10)
            {
                bool isLetter = false;
                bool isDigit = false;
                bool isSpecialChar = false;

                foreach (char letter in txtPasswordSignUp.Text)
                {
                    if (char.IsLetter(letter))
                    {
                        isLetter = true;
                    }
                    else if (char.IsDigit(letter))
                    {
                        isDigit = true;
                    }
                    else if (char.IsSymbol(letter) || char.IsPunctuation(letter))
                    {
                        isSpecialChar = true;
                    }

                    if (isLetter && isDigit && isSpecialChar)
                    {
                        break;
                    }
                }

                bool goodPassword = isLetter && isDigit && isSpecialChar;
                if (!goodPassword)
                {
                    if(!isLetter)
                        MessageBox.Show("The password must contains at least one letter");
                    if(!isDigit)
                        MessageBox.Show("The password must contains at least one digit");
                    if(!isSpecialChar)
                        MessageBox.Show("The password must contains at least one special character");
                    flag = false;
                }
            }
            else
            {
                MessageBox.Show("The password must contain between 8 and 10 characters");
                flag = false;
            }

            if (flag)
            {
                string exelNameUsers = "ProjectUsers.xlsx";
                string fileName = Directory.GetCurrentDirectory() + "\\" + exelNameUsers;
                User newUser = new User(txtUserNameSignUp.Text, txtPasswordSignUp.Text, IDSingUp.Text);
                WriteToExcel(fileName, txtUserNameSignUp.Text, txtPasswordSignUp.Text, IDSingUp.Text);
                MainWindow windowM = new MainWindow();
                windowM.Show();
                this.Close();
            }


        }


        public static void WriteToExcel(string filePath, string userName, string password, string id)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            // Find the first empty row to write the data
            int newRow = rowCount + 1;

            // Write the username, password, and ID to the worksheet
            usedRange.Cells[newRow, 1].Value = userName; // Assuming name is in column 1
            usedRange.Cells[newRow, 2].Value = password; // Assuming password is in column 2
            usedRange.Cells[newRow, 3].Value = id; // Assuming ID is in column 3

            workbook.Save();
            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
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


        public static Boolean SearchExcel(string filePath, string UserName)
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

                if (nameInCell == UserName)
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


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }
    }
}
