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
    /// Interaction logic for AddBird.xaml
    /// </summary>
    public partial class AddBird : System.Windows.Window
    {
        public AddBird()
        {
            InitializeComponent();
            Loaded += AddBird_Loaded;
        }

        private void AddBird_Loaded(object sender, RoutedEventArgs e)
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



        private void cmbSpecies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cmbSubSpecies.Items.Clear();
            if (cmbSpecies.SelectedItem is ComboBoxItem selectedItem)
            {
                string species = selectedItem.Content.ToString();
                switch (species)
                {
                    case "American Gouldian":
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "North America" });
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "Center America" });
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "South America" });
                        break;
                    case "European Gouldian":
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "East Europe" });
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "West Europe" });
                        break;
                    case "Australian Gouldian":
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "Center Australia" });
                        cmbSubSpecies.Items.Add(new ComboBoxItem() { Content = "Coastal Cities" });
                        break;
                    default:
                      
                        break;
                }
            }
        }

        private void btnAddBird_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;
            bool isLetter = false;
            bool isDigit = false;
            string SerialMom ="";
            string SerialDad ="";

            if (!ChackSerial(txtSerialNumber.Text)) {
                MessageBox.Show("The serial number must contain only digits");
                flag = false;
            }

            string exelNameBird = "Birds.xlsx";
            string fileBird = Directory.GetCurrentDirectory() + "\\" + exelNameBird;

            if (SearchExcel(fileBird, txtSerialNumber.Text,1))
            {
                MessageBox.Show("This serial number alrady exist");
                flag = false;
            }

            if (gridDate.SelectedDate.ToString() == "")
            {
                MessageBox.Show("You must choose an hatch date");
                flag = false;
            }

            if (gridDate.SelectedDate > DateTime.Now)
            {
                MessageBox.Show("The inputted DateTime is later than the current time");
                flag = false;
            }

            if (cmbSpecies.Text == "")
            {
                MessageBox.Show("You must choose a species");
                flag = false;
            }

            if (cmbSubSpecies.Text == "")
            {
                MessageBox.Show("You must choose a subspecies");
                flag = false;
            }

            if (cmbGender.Text == "")
            {
                MessageBox.Show("You must choose a gender");
                flag = false;
            }

            foreach (char letter in txtNumberCage.Text)
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
                MessageBox.Show("The cage must contains at least one letter, one digit");
                flag = false;
            }

            string exelNameCage = "cages.xlsx";
            string fileCage = Directory.GetCurrentDirectory() + "\\" + exelNameCage;

            if (!SearchExcel(fileCage, txtNumberCage.Text,1))
            {
                MessageBox.Show("The cage is not exist");
                flag = false;
            }

            string genderBirdDad = SearchExcelGender(fileBird, txtSerialDad.Text, 1);
            string genderBirdMom = SearchExcelGender(fileBird, txtSerialMom.Text, 1);
            

            if (txtSerialDad.Text == "" || txtSerialMom.Text == "")
            {
                if(txtSerialDad.Text == "" && txtSerialMom.Text != "")
                {
                    if (!ChackSerial(txtSerialMom.Text))
                    {
                        MessageBox.Show("The mother serial number must contain only digits");
                        flag = false;
                    }
                    else
                    {
                        if (genderBirdMom == "Male")
                        {
                            MessageBox.Show("The mother serial number must be of a female Bird");
                            flag = false;
                        }
                        else
                            SerialMom = txtSerialMom.Text;
                    }
                    SerialDad = "0";
                }
                if (txtSerialMom.Text == "" && txtSerialDad.Text != "")
                {
                    if (!ChackSerial(txtSerialDad.Text))
                    {
                        MessageBox.Show("The father serial number must contain only digits");
                        flag = false;
                    }
                    else
                    {
                        if (genderBirdDad == "Female")
                        {
                            MessageBox.Show("The father serial number must be of a male Bird");
                            flag = false;
                        }
                        else
                            SerialDad = txtSerialDad.Text;
                    }
                    SerialMom = "0";
                }

            }

            if (txtSerialDad.Text == "" && txtSerialMom.Text == "")     //A bird without a mother and father
            {
                SerialDad = "0";
                SerialMom = "0";
            }


            if (txtSerialDad.Text != "" && txtSerialMom.Text != "")
            {
                if (!ChackSerial(txtSerialDad.Text))
                {
                    MessageBox.Show("The father serial number must contain only digits");
                    flag = false;
                }

                if (genderBirdDad == "Female")
                {
                    MessageBox.Show("The father serial number must be of a male Bird");
                    flag = false;
                }
                else
                {
                    SerialDad = txtSerialDad.Text;
                }


                if (!ChackSerial(txtSerialMom.Text))
                {
                    MessageBox.Show("The mother serial number must contain only digits");
                    flag = false;
                }

                if (genderBirdMom == "Male")
                {
                    MessageBox.Show("The mother serial number must be of a female Bird");
                    flag = false;
                }
                else
                {
                    SerialMom = txtSerialMom.Text;
                }
            }

            if (flag)
            {
                Bird newBird = new Bird(int.Parse(txtSerialNumber.Text), cmbSpecies.Text, cmbSubSpecies.Text, gridDate.DisplayDate, cmbGender.Text, txtNumberCage.Text, SerialMom, SerialDad);
                WriteToExcel(fileBird, txtSerialNumber.Text, cmbSpecies.Text, cmbSubSpecies.Text, gridDate.ToString(), cmbGender.Text, txtNumberCage.Text, SerialMom, SerialDad);
                MessageBox.Show("The Bird was added successfully");
                AfterLogin afterLoginWindow = new AfterLogin();
                afterLoginWindow.Show();
                this.Close();
            }
        }

        public Boolean ChackSerial(string s)
        {
            if (s == "")
                return false;

            for(int i=0; i < s.Length; i++)
            {
                if (!char.IsDigit(s[i]))
                    return false;
            }
            return true;
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
                string Cell = usedRange.Cells[row, c].Value?.ToString();
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

        public static void WriteToExcel(string filePath, string serialNumber, string species, string subspecies, string date, string gender, string cageNumber, string motherSerial, string fatherSerial)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            // Find the first empty row to write the data
            int newRow = rowCount + 1;

            DateTime hatchDate = DateTime.Parse(date);
            string formattedDate = hatchDate.ToString("MM/dd/yyyy");


            // Write the data to the worksheet
            usedRange.Cells[newRow, 1].Value = serialNumber; // Serial number
            usedRange.Cells[newRow, 2].Value = species; // Species
            usedRange.Cells[newRow, 3].Value = subspecies; // Subspecies
            usedRange.Cells[newRow, 4].Value = formattedDate; // Date
            usedRange.Cells[newRow, 5].Value = gender; // Gender
            usedRange.Cells[newRow, 6].Value = cageNumber; // Number of cage
            usedRange.Cells[newRow, 7].Value = motherSerial; // Serial number of mother
            usedRange.Cells[newRow, 8].Value = fatherSerial; // Serial number of father

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

        public string SearchExcelGender(string filePath, string Value, int c)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;
            string gender = "";
            for (int row = 1; row <= rowCount; row++)
            {
                string Cell = usedRange.Cells[row, c].Value?.ToString();
                if (Cell == Value)
                {
                    gender = usedRange.Cells[row, 5].Value?.ToString();
                    workbook.Close();
                    excel.Quit();
                    ReleaseObject(worksheet);
                    ReleaseObject(workbook);
                    ReleaseObject(excel);
                    return gender;
                }
            }

            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
            return gender;
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
