using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace BirdsProject1
{
    /// <summary>
    /// Interaction logic for EditCage.xaml
    /// </summary>
    public partial class EditCage : System.Windows.Window
    {
        static string exelNameCage = "cages.xlsx";
        string fileCage = Directory.GetCurrentDirectory() + "\\" + exelNameCage;

        static string exelNameBird = "Birds.xlsx";
        string fileBird = Directory.GetCurrentDirectory() + "\\" + exelNameBird;

        public Cage OldCage;

        private void EditCage_Loaded(object sender, RoutedEventArgs e)
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

        public EditCage(Cage newCage)
        {
            InitializeComponent();
            Loaded += EditCage_Loaded;
            OldCage = newCage;
            txtSerialnumberCage.Text = newCage.SerialNumber;
            cmbMaterial.Text = newCage.Material;
            txtLengthCage.Text = newCage.Length.ToString();
            txtHeightCage.Text = newCage.Height.ToString();
            txtWidthCage.Text = newCage.Width.ToString();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            bool isLetter = false;
            bool isDigit = false;

            foreach (char letter in txtSerialnumberCage.Text)
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
                return;
            }

            if (!check_validation(txtLengthCage.Text, "length"))
                return;
            if (!check_validation(txtWidthCage.Text, "width"))
                return;
            if (!check_validation(txtHeightCage.Text, "height"))
                return;
            if (int.Parse(txtHeightCage.Text) > 3000)
            {
                MessageBox.Show("The height must be under 3001 meters");
                return;
            }

            if (cmbMaterial.Text == "")
            {
                MessageBox.Show("You must choose a material");
                return;
            }

            if (txtSerialnumberCage.Text != OldCage.SerialNumber)
            {
                if (HaveACage(txtSerialnumberCage.Text))
                {
                    MessageBox.Show("This Cage exists! Please enter the correct number of the cage.");
                    return;
                }

                EditExcelRow(txtSerialnumberCage.Text, txtLengthCage.Text, txtWidthCage.Text, txtHeightCage.Text, cmbMaterial.Text);
                DeleteAndUpdateRow(OldCage.SerialNumber, txtSerialnumberCage.Text);
                MessageBox.Show("The cage has been updated");
                Close();
            }
            else
            {
                EditExcelRow(txtSerialnumberCage.Text, txtLengthCage.Text, txtWidthCage.Text, txtHeightCage.Text, cmbMaterial.Text);
                MessageBox.Show("The cage has been updated");
                Close();
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

        public void EditExcelRow(string serialNumberCage, string LengthCage, string WidthCage, string HeightCage, string Material)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excel.Workbooks.Open(fileCage);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            // Find the row to edit
            for (int row = 2; row <= rowCount; row++)
            {
                string serialNumber = usedRange.Cells[row, 1].Value?.ToString();

                if (serialNumber == OldCage.SerialNumber)
                {
                    usedRange.Cells[row, 1].Value = serialNumberCage; // Serial number
                    usedRange.Cells[row, 2].Value = LengthCage; // Length
                    usedRange.Cells[row, 3].Value = WidthCage; // Width
                    usedRange.Cells[row, 4].Value = HeightCage; // Height
                    usedRange.Cells[row, 5].Value = Material; // Material
                    break;
                }

            }

            workbook.Save();
            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);

        }

        private void ReleaseObject(object obj)
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

        public bool HaveACage(string value)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excel.Workbooks.Open(fileCage);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            for (int row = 2; row <= rowCount; row++)
            {
                string serialNumber = usedRange.Cells[row, 1].Value?.ToString();

                if (serialNumber == value)
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

        private void DeleteAndUpdateRow(string oldCageNumber, string newCageNumber)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(fileBird);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;

            for (int row = rowCount; row >= 2; row--)
            {
                string cageNumber = usedRange.Cells[row, 6].Value?.ToString();

                if (cageNumber == oldCageNumber)
                {
                    string SerialNumberb = usedRange.Cells[row, 1].Value?.ToString();
                    string species = usedRange.Cells[row, 2].Value?.ToString();
                    string subSpecies = usedRange.Cells[row, 3].Value?.ToString();
                    string hatchingDate = usedRange.Cells[row, 4].Value?.ToString();
                    string gender = usedRange.Cells[row, 5].Value?.ToString();
                    string serialNumberMother = usedRange.Cells[row, 7]?.Value.ToString();
                    string serialNumberFather = usedRange.Cells[row, 8]?.Value.ToString();

                    // Write the new row with updated information
                    int newRow = row;
                    usedRange.Cells[newRow, 1].Value = SerialNumberb;
                    usedRange.Cells[newRow, 2].Value = species;
                    usedRange.Cells[newRow, 3].Value = subSpecies;
                    usedRange.Cells[newRow, 4].Value = hatchingDate;
                    usedRange.Cells[newRow, 5].Value = gender;
                    usedRange.Cells[newRow, 6].Value = newCageNumber;
                    usedRange.Cells[newRow, 7].Value = serialNumberMother;
                    usedRange.Cells[newRow, 8].Value = serialNumberFather;

                    // Optionally, you can store the updated bird's information if necessary:
                    Bird bird = new Bird(
                        int.Parse(SerialNumberb),
                        species,
                        subSpecies,
                        DateTime.Parse(hatchingDate),
                        gender,
                        newCageNumber,
                        serialNumberMother,
                        serialNumberFather
                    );
                }
            }

            workbook.Save();
            workbook.Close();
            excel.Quit();
            ReleaseObject(worksheet);
            ReleaseObject(workbook);
            ReleaseObject(excel);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }
    }
}