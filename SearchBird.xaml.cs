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
    /// Interaction logic for SearchBird.xaml
    /// </summary>
    public partial class SearchBird : System.Windows.Window
    {
        static string exelNameBird = "Birds.xlsx";
        string filePath = Directory.GetCurrentDirectory() + "\\" + exelNameBird;

        private List<Bird> matchingRows;
        public SearchBird()
        {
            InitializeComponent();
            Loaded += SearchBird_Loaded;
            DataContext = this;
            dgAllTheFoundBirds.MouseDoubleClick += DgAllTheFoundBirds_MouseDoubleClick;
        }

        private void SearchBird_Loaded(object sender, RoutedEventArgs e)
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

        private void DgAllTheFoundBirds_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Retrieve the selected bird from the data grid
            Bird selectedBird = dgAllTheFoundBirds.SelectedItem as Bird;
            if (selectedBird != null)
            {
                // Open a new window and pass the selected bird to it
                DisplayBird birdDetailsWindow = new DisplayBird(selectedBird);
                birdDetailsWindow.Show();
            }
        }


        private void MainComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Hide all additional controls
            additionalComboBoxGrid.Visibility = Visibility.Collapsed;
            MaleOrFemaleComboBox.Visibility = Visibility.Collapsed;
            SerialNumberTextBox.Visibility = Visibility.Collapsed;
            SpeciesComboBox.Visibility = Visibility.Collapsed;
            DatePickerDateLabel.Visibility = Visibility.Collapsed;

            // Check the selected item in the mainComboBox
            ComboBoxItem selectedItem = mainComboBox.SelectedItem as ComboBoxItem;
            if (selectedItem != null)
            {
                string selectedItemContent = selectedItem.Content.ToString();
                if (selectedItemContent == "Serial number")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    SerialNumberTextBox.Visibility = Visibility.Visible;
                }
                else if (selectedItemContent == "Gender")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    MaleOrFemaleComboBox.Visibility = Visibility.Visible;
                }
                else if (selectedItemContent == "Species")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    SpeciesComboBox.Visibility = Visibility.Visible;
                }
                else if (selectedItemContent == "Hatch date")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    DatePickerDateLabel.Visibility = Visibility.Visible;
                }
            }
        }

        private void btnSeachBird_Click(object sender, RoutedEventArgs e)
        {
            if (SerialNumberTextBox.Text != "") {
                SearchExcel(SerialNumberTextBox.Text, 1);
            }
            else if (MaleOrFemaleComboBox.Text != "") {
                string x = MaleOrFemaleComboBox.Text;
                SearchExcel(x, 5);
            }
            else if (SpeciesComboBox.Text != "") {
                SearchExcel(SpeciesComboBox.Text, 2);
            }
            else if (DatePickerDateLabel.Text != "") {
                string value = DatePickerDateLabel.Text + " 00:00:00";
                SearchExcel(value, 4);
            }
            else
            {
                MessageBox.Show("You must select a search variable");
            }
            SerialNumberTextBox.Text = "";
            MaleOrFemaleComboBox.Text = "";
            SpeciesComboBox.Text = "";
            DatePickerDateLabel.Text = "";
        }

        private void SearchExcel(string value, int c)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            matchingRows = new List<Bird>();
            
            for (int row=2; row <= rowCount; row++) // Skip header row
            {
                string SerialNumberb = usedRange.Cells[row, 1].Value?.ToString();
                string species = usedRange.Cells[row, 2].Value?.ToString();
                string subSpecies = usedRange.Cells[row, 3].Value?.ToString();
                string hatchingDate = usedRange.Cells[row, 4].Value?.ToString();
                string gender = usedRange.Cells[row, 5].Value?.ToString();
                string cageNumber = usedRange.Cells[row, 6].Value?.ToString();
                string serialNumberMother = usedRange.Cells[row, 7]?.Value.ToString();
                string serialNumberFather = usedRange.Cells[row, 8]?.Value.ToString();


                if (gender == value || SerialNumberb == value || species == value || hatchingDate == value)
                {
                    Bird bird = new Bird(int.Parse(SerialNumberb), species, subSpecies, DateTime.Parse(hatchingDate), gender, cageNumber, serialNumberMother, serialNumberFather);
                    matchingRows.Add(bird);

                }
            }
            if (matchingRows.Count == 1)
            {
                // Open a new window and pass the selected bird to it
                DisplayBird birdDetailsWindow = new DisplayBird(matchingRows[0]);
                birdDetailsWindow.Show();
                //this.Close();

                workbook.Close();
                excel.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excel);
            }
            else      //There is more then 1 bird in the list - display the deatils in the dataGrid
            {
                matchingRows.Sort((x, y) => x.SerialNumber.CompareTo(y.SerialNumber));    //Sort the list
                dgAllTheFoundBirds.ItemsSource = matchingRows.ToList();    //dispaly the data for the birds in the data grid

                workbook.Close();
                excel.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excel);
            }
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }
    }

}