using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Data.SqlClient;
using System.Windows.Shapes;

namespace WpfApp9
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //await LoadCategoriesAsync();
        }
        public MainWindow()
        {
            InitializeComponent();
        }
        private void OpenSecondWindowButton_Click(object sender, RoutedEventArgs e)
        {
            Window1 window1 = new Window1();
            window1.ShowDialog(); 

        }
        private void OpenRegWindowButton_Click(object sender, RoutedEventArgs e)
        {
            Reg reg = new Reg();
            reg.ShowDialog();
            
        }
    }
}