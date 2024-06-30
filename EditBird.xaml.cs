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
    public partial class EditBird : System.Windows.Window
    {
        private Bird selectedBird;
        private Bird originalBird;
        static string exelNameBird = "Birds.xlsx";
        string fileBird = Directory.GetCurrentDirectory() + "\\" + exelNameBird;


        private void EditBird_Loaded(object sender, RoutedEventArgs e)
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

        public EditBird(Bird bird)
        {
            InitializeComponent();
            Loaded += EditBird_Loaded;
            selectedBird = bird;
            originalBird = bird;
            DataContext = selectedBird;
            txtSerialNumber.Text = bird.SerialNumber.ToString();
            cmbSpecies.Text = bird.species;
            cmbSubSpecies.Text = bird.subSpecies;
            gridDate.Text = bird.hatchDate.ToString();
            cmbGender.Text = bird.gender;
            txtNumberCage.Text = bird.cageNumber;
            txtSerialMom.Text = bird.SerialNumberMother;
            txtSerialDad.Text = bird.SerialNumberfather;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;

            if (!ChackSerial(txtSerialNumber.Text))
            {
                MessageBox.Show("The serial number must contain only digits");
                flag = false;
            }

            if (txtSerialNumber.Text != originalBird.SerialNumber.ToString())
            {
                if (HaveABird(txtSerialNumber.Text))
                {
                    MessageBox.Show("This serial number exists! Please enter the correct serial number of the bird.");
                    flag = false;
                }
            }

            if (gridDate.SelectedDate > DateTime.Now)
            {
                MessageBox.Show("The inputted DateTime is later than the current time");
                flag = false;
            }

            string exelNameCage = "cages.xlsx";
            string fileCage = Directory.GetCurrentDirectory() + "\\" + exelNameCage;

            if (!SearchExcel(fileCage, txtNumberCage.Text, 1))
            {
                MessageBox.Show("The cage is not exist");
                flag = false;
            }

            if (txtSerialDad.Text != "" || txtSerialMom.Text != "")
            {
                if (!ChackSerial(txtSerialDad.Text))
                {
                    MessageBox.Show("The father serial number must contain only digits");
                    flag = false;
                }
                if (!ChackSerial(txtSerialMom.Text))
                {
                    MessageBox.Show("The mother serial number must contain only digits");
                    flag = false;
                }
            }

            if (flag)
            {
                EditExcelRow(txtSerialNumber.Text, cmbSpecies.Text, cmbSubSpecies.Text, gridDate.ToString(), cmbGender.Text, txtNumberCage.Text, txtSerialMom.Text, txtSerialDad.Text);
                MessageBox.Show("The Bird has been updated");
                this.Close();
            }

            
        }

        public bool HaveABird(string value)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excel.Workbooks.Open(fileBird);
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
        public Boolean ChackSerial(string s)
        {
            if (s == "")
                return false;

            for (int i = 0; i < s.Length; i++)
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

        public void EditExcelRow(string serialNumberbird, string species, string subspecies, string hatchDate, string gender, string cageNumber, string SerialNumberMother, string SerialNumberfather)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excel.Workbooks.Open(fileBird);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            // Find the row to edit
            for (int row = 2; row <= rowCount; row++)
            {
                string serialNumber = usedRange.Cells[row, 1].Value?.ToString();

                if (serialNumber == originalBird.SerialNumber.ToString())
                {
                    DateTime HDate = DateTime.Parse(hatchDate);
                    string formattedDate = HDate.ToString("MM/dd/yyyy");

                    usedRange.Cells[row, 1].Value = serialNumberbird; // Serial number
                    usedRange.Cells[row, 2].Value = species; 
                    usedRange.Cells[row, 3].Value = subspecies; 
                    usedRange.Cells[row, 4].Value = formattedDate; 
                    usedRange.Cells[row, 5].Value = gender; 
                    usedRange.Cells[row, 6].Value = cageNumber;
                    usedRange.Cells[row, 7].Value = SerialNumberMother;
                    usedRange.Cells[row, 8].Value = SerialNumberfather;
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








