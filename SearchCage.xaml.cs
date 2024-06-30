using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;/////////
using System.IO;

namespace BirdsProject1
{
    /// <summary>
    /// Interaction logic for SearchCage.xaml
    /// </summary>
    public partial class SearchCage : System.Windows.Window
    {
        static string exelNameCage = "cages.xlsx";
        string filePath = Directory.GetCurrentDirectory() + "\\" + exelNameCage;

        private List<Cage> searchResults;

        private void SearchCage_Loaded(object sender, RoutedEventArgs e)
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

        public SearchCage()
        {
            InitializeComponent();
            Loaded += SearchCage_Loaded;
            DataContext = this;
            dataGridCage.MouseDoubleClick += DataGridCage_MouseDoubleClick;
        }

        private void cmbSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Hide all additional controls
            additionalComboBoxGrid.Visibility = Visibility.Collapsed;
            MaterialComboBox.Visibility = Visibility.Collapsed;
            SerialNumberTextBox.Visibility = Visibility.Collapsed;

            // Check the selected item in the mainComboBox
            ComboBoxItem selectedItem = cmbSelect.SelectedItem as ComboBoxItem;
            if (selectedItem != null)
            {
                string selectedItemContent = selectedItem.Content.ToString();
                if (selectedItemContent == "Serial Number")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    SerialNumberTextBox.Visibility = Visibility.Visible;
                }
                else if (selectedItemContent == "Material")
                {
                    additionalComboBoxGrid.Visibility = Visibility.Visible;
                    MaterialComboBox.Visibility = Visibility.Visible;
                }
            }
        }

        private void SearchExcel(string value, int columnIndex)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first worksheet

            Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int columnCount = usedRange.Columns.Count;

            searchResults = new List<Cage>();

            for (int row = 2; row <= rowCount; row++)
            {
                string serialNumber = usedRange.Cells[row, 1].Value?.ToString();
                string length = usedRange.Cells[row, 2].Value?.ToString();
                string width = usedRange.Cells[row, 3].Value?.ToString();
                string height = usedRange.Cells[row, 4].Value?.ToString();
                string material = usedRange.Cells[row, 5].Value?.ToString();

                if (serialNumber == value || material == value)
                {
                    Cage cage = new Cage(serialNumber, int.Parse(length), int.Parse(width), int.Parse(height), material);
                    searchResults.Add(cage);

                }
            }
            dataGridCage.ItemsSource = searchResults.ToList();

            if (searchResults.Count == 1)
            {
                DisplayCage displayCageWindow = new DisplayCage(searchResults[0]);
                displayCageWindow.Show();
                this.Close();

                workbook.Close();
                excel.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excel);

            }
            // Display the search results in the datagrid
            else
            {
                workbook.Close();
                excel.Quit();
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excel);
            }
        }
        private void DataGridCage_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Retrieve the selected bird from the data grid
            Cage selectedCage = dataGridCage.SelectedItem as Cage;

            if (selectedCage != null)
            {
                // Open a new window and pass the selected bird to it
                DisplayCage birdDetailsWindow = new DisplayCage(selectedCage);
                birdDetailsWindow.Show();
                //this.Close();
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

        private void btnSearchCage_Click(object sender, RoutedEventArgs e)
        {
            bool flag = true;
            bool isLetter = false;
            bool isDigit = false;
            if (SerialNumberTextBox.Text != "") {
                foreach (char letter in SerialNumberTextBox.Text)
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
                if (flag)
                    SearchExcel(SerialNumberTextBox.Text, 1);
            }

            else if (MaterialComboBox.Text != "")
                SearchExcel(MaterialComboBox.Text, 3);
            else
                MessageBox.Show("You must select a search variable");
            SerialNumberTextBox.Text = "";
            MaterialComboBox.Text = "";
        }

        private void btnAfterLogin_Click(object sender, RoutedEventArgs e)
        {
            AfterLogin afterLoginWindow = new AfterLogin();
            afterLoginWindow.Show();
            this.Close();
        }

        private void btnMainwindow_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}