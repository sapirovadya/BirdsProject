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
    /// Interaction logic for AddFledgling.xaml
    /// </summary>
    public partial class AddFledgling : System.Windows.Window
    {
        private Bird parentBird;
        private Bird parentBirdSecond;
        private string MomOrDad;

        private void AddFledgling_Loaded(object sender, RoutedEventArgs e)
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

        public AddFledgling(Bird bird)
        {
            InitializeComponent();
            Loaded += AddFledgling_Loaded;
            parentBird = bird;
            DataContext = parentBird;
            txtSerialNumber.Text = "";
            txtSpecies.Text = bird.species;
            txtSubSpecies.Text = bird.subSpecies;
            txtNumberCage.Text = bird.cageNumber;
            txtSerialKnowParent.Text = bird.SerialNumber.ToString();
            txtSerialSecondParent.Text = "";


            if (bird.gender == "Female")
            {
                MomOrDad = "mom";
                labelSerialMom.Visibility = Visibility.Hidden;
                labelSerialDad.Visibility = Visibility.Visible;

                labelSerialMom_Know.Visibility = Visibility.Visible;
                labelSerialDad_Know.Visibility = Visibility.Hidden;
            }
            else
            {
                MomOrDad = "dad";
                labelSerialMom.Visibility = Visibility.Visible;
                labelSerialDad.Visibility = Visibility.Hidden;

                labelSerialMom_Know.Visibility = Visibility.Hidden;
                labelSerialDad_Know.Visibility = Visibility.Visible;
            }
            
    }

        private void btnAddFledgling_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;

            int serialNumber;
            if (!int.TryParse(txtSerialNumber.Text, out serialNumber))
            {
                MessageBox.Show("The serial number must contain only digits");
                flag = false;
            }

            string exelNameBird = "Birds.xlsx";
            string fileBird = Directory.GetCurrentDirectory() + "\\" + exelNameBird;

            if (SearchExcel(fileBird, txtSerialNumber.Text, 1))
            {
                MessageBox.Show("This serial number already exsist");
                flag = false;
            }

            if (txtSerialSecondParent.Text != "" || txtSerialSecondParent.Text != " ")
            {
                if (!SearchExcel(fileBird, txtSerialSecondParent.Text, 1))
                {
                    MessageBox.Show("This serial number of the second parent is not exsist");
                    flag = false;
                }
            }
            else
            {
                MessageBox.Show("The serial number of the other parent must be entered");
                flag = false;
            }

            parentBirdSecond = SearchExcelbuildBird(fileBird, txtSerialSecondParent.Text, 1);

            if (parentBird.gender == "Female" && parentBirdSecond.gender == "Female")
            {
                MessageBox.Show("The gender of the other parent must be male, please choose serial number of male bird");
                flag = false;
            }
            else if (parentBird.gender == "Male" && parentBirdSecond.gender == "Male")
            {
                MessageBox.Show("The gender of the other parent must be female, please choose serial number of female bird");
                flag = false;
            }

            if (gridDate.Text == "")
            {
                MessageBox.Show("You must enter a hatch date");
                flag = false;
            }

            if (gridDate.SelectedDate > DateTime.Now)
            {
                MessageBox.Show("The inputted hatch date is later than the current time");
                flag = false;
            }
            else if (gridDate.SelectedDate <= parentBird.hatchDate || gridDate.SelectedDate <= parentBirdSecond.hatchDate)
            {
                MessageBox.Show("The inputted hatch date is earlier than the parent");
                flag = false;
            }

            if(cmbGender.Text == "" || cmbGender.Text == " ")
            {
                MessageBox.Show("You nust enter a gender");
                flag = false;
            }

            if (flag)
            {
                Bird newFledgling = null;
                if (MomOrDad == "mom")
                {
                    if (SearchExcel(fileBird, txtSerialSecondParent.Text, 1))
                    {
                        newFledgling = new Bird(int.Parse(txtSerialNumber.Text), txtSpecies.Text, txtSubSpecies.Text, gridDate.DisplayDate, cmbGender.Text, txtNumberCage.Text, txtSerialKnowParent.Text, txtSerialSecondParent.Text);
                        WriteToExcel(fileBird, txtSerialNumber.Text, txtSpecies.Text, txtSubSpecies.Text, gridDate.ToString(), cmbGender.Text, txtNumberCage.Text, txtSerialKnowParent.Text, txtSerialSecondParent.Text);
                    }
                }
                else
                {
                    newFledgling = new Bird(int.Parse(txtSerialNumber.Text), txtSpecies.Text, txtSubSpecies.Text, gridDate.DisplayDate, cmbGender.Text, txtNumberCage.Text, txtSerialSecondParent.Text, txtSerialKnowParent.Text);
                    WriteToExcel(fileBird, txtSerialNumber.Text, txtSpecies.Text, txtSubSpecies.Text, gridDate.ToString(), cmbGender.Text, txtNumberCage.Text, txtSerialSecondParent.Text, txtSerialKnowParent.Text);

                }

                MessageBox.Show("The Fledgling was added successfully");
                this.Close();
            }

        }

        public static bool SearchExcel(string filePath, string Value, int col)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            for (int row = 2; row <= rowCount; row++)
            {
                string Cell = usedRange.Cells[row, col].Value?.ToString();
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

        public Bird SearchExcelbuildBird(string filePath, string Value, int col)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            for (int row = 2; row <= rowCount; row++)
            {
                string Cell = usedRange.Cells[row, col].Value?.ToString();
                if (Cell == Value)
                {
                    string SerialNumberb = usedRange.Cells[row, 1].Value?.ToString();
                    string species = usedRange.Cells[row, 2].Value?.ToString();
                    string subSpecies = usedRange.Cells[row, 3].Value?.ToString();
                    string hatchingDate = usedRange.Cells[row, 4].Value?.ToString();
                    string gender = usedRange.Cells[row, 5].Value?.ToString();
                    string cageNumber = usedRange.Cells[row, 6].Value?.ToString();
                    string serialNumberMother = usedRange.Cells[row, 7]?.Value.ToString();
                    string serialNumberFather = usedRange.Cells[row, 8]?.Value.ToString();

                    parentBirdSecond = new Bird(int.Parse(SerialNumberb), species, subSpecies, DateTime.Parse(hatchingDate), gender, cageNumber, serialNumberMother, serialNumberFather);
                }
            }
            
            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
            return parentBirdSecond;
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

            // Write the data to the worksheet
            usedRange.Cells[newRow, 1].Value = serialNumber; // Serial number
            usedRange.Cells[newRow, 2].Value = species; // Species
            usedRange.Cells[newRow, 3].Value = subspecies; // Subspecies
            usedRange.Cells[newRow, 4].Value = date; // Date
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }
    }
}
