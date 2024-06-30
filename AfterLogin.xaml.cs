using System;
using System.Collections.Generic;
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

namespace BirdsProject1
{
    /// <summary>
    /// Interaction logic for AfterLogin.xaml
    /// </summary>
    public partial class AfterLogin : Window
    {
        public AfterLogin()
        {
            InitializeComponent();
            Loaded += AfterLogin_Loaded;
        }

        private void AfterLogin_Loaded(object sender, RoutedEventArgs e)
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

        private void btnAddBird_Click(object sender, RoutedEventArgs e)
        {
            AddBird addBirdWindow = new AddBird();
            addBirdWindow.Show();
            this.Close();
        }

        private void btnAddCage_Click(object sender, RoutedEventArgs e)
        {
            AddCage addCageWindow = new AddCage();
            addCageWindow.Show();
            this.Close();

        }

        private void btnSearchCage_Click(object sender, RoutedEventArgs e)
        {
            SearchCage searchCageWindow = new SearchCage();
            searchCageWindow.Show();
            this.Close();
        }

        private void btnSearchBird_Click(object sender, RoutedEventArgs e)
        {
            SearchBird searchBird = new SearchBird();
            searchBird.Show();
            this.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
