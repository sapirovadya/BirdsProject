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
    /// Interaction logic for DisplayCage.xaml
    /// </summary>
    public partial class DisplayCage : System.Windows.Window
    {
        public List<Bird> ListBirds = new List<Bird>();
        public Cage OldCage;

        private void DisplayCage_Loaded(object sender, RoutedEventArgs e)
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

        public DisplayCage(Cage newCage)
        {
            InitializeComponent();
            Loaded += DisplayCage_Loaded;
            OldCage = newCage;
            txtSerialNumberCage.Text = newCage.SerialNumber;
            txtMaterial.Text = newCage.Material;
            txtLengthCage.Text = newCage.Length.ToString();
            txtHeightCage.Text = newCage.Height.ToString();
            txtWidthCage.Text = newCage.Width.ToString();
            SearchExcel(newCage.SerialNumber);
        }

        private void btnAddCage_Click(object sender, RoutedEventArgs e)
        {
            EditCage editCageWindow = new EditCage(OldCage);
            editCageWindow.ShowDialog();
            this.Close();
        }

        private void SearchExcel(string value)
        {
            string exelNameBird = "Birds.xlsx";
            string filePath = Directory.GetCurrentDirectory() + "\\" + exelNameBird;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            ListBirds = new List<Bird>();

            for (int row = 2; row <= rowCount; row++) // Skip header row
            {
                string SerialNumberb = usedRange.Cells[row, 1].Value?.ToString();
                string species = usedRange.Cells[row, 2].Value?.ToString();
                string subSpecies = usedRange.Cells[row, 3].Value?.ToString();
                string hatchingDate = usedRange.Cells[row, 4].Value?.ToString();
                string gender = usedRange.Cells[row, 5].Value?.ToString();
                string cageNumber = usedRange.Cells[row, 6].Value?.ToString();
                string serialNumberMother = usedRange.Cells[row, 7]?.Value.ToString();
                string serialNumberFather = usedRange.Cells[row, 8]?.Value.ToString();

                Console.WriteLine("in the faile: " + hatchingDate);

                if (cageNumber == value)
                {
                    Bird bird = new Bird(int.Parse(SerialNumberb), species, subSpecies, DateTime.Parse(hatchingDate), gender, cageNumber, serialNumberMother, serialNumberFather);
                    ListBirds.Add(bird);

                }
            }

            datagridAllBird.ItemsSource = ListBirds.ToList();

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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }
    }
}