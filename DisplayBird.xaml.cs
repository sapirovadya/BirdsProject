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
    /// Interaction logic for DisplayBird.xaml
    /// </summary>
    public partial class DisplayBird : Window
    {
        private Bird selectedBird;

        private void DisplayBird_Loaded(object sender, RoutedEventArgs e)
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
        public DisplayBird(Bird bird)
        {
            InitializeComponent();
            Loaded += DisplayBird_Loaded;
            selectedBird = bird;
            DataContext = selectedBird;
            txtSerialNumber.Text = bird.SerialNumber.ToString();
            txtSpecies.Text = bird.species;
            txtSubSpecies.Text = bird.subSpecies;
            txtgridDate.Text = bird.hatchDate.ToString();
            txtGender.Text = bird.gender;
            txtNumberCage.Text = bird.cageNumber;
            txtSerialMom.Text = bird.SerialNumberMother;
            txtSerialDad.Text = bird.SerialNumberfather;
        }

        private void btnEditBird_Click(object sender, RoutedEventArgs e)
        {
            EditBird birdDetailsWindow = new EditBird(selectedBird);
            birdDetailsWindow.Show();
            this.Close();
        }

        private void btnAddFledgling_Click(object sender, RoutedEventArgs e)
        {
            AddFledgling newFledgling = new AddFledgling(selectedBird);
            newFledgling.Show();
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
