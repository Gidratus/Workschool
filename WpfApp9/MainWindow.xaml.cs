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
        private const string ConnectionString =
           "Server=localhost\\SQLEXPRESS;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";
        private const string Sql = "select * from dbo.Categories";
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //await LoadCategoriesAsync();
        }
        public MainWindow()
        {
            InitializeComponent();
        }
        private async void OpenSecondWindowButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(LastNameTextBox.Text))
            {
                MessageBox.Show("введите фамилию!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(FirstnameTextbox.Text))
            {
                MessageBox.Show("введите Имя!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(PasswordTextbox.Password))
            {
                MessageBox.Show("введите Пароль!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            bool isAuthenticated = await AuthenticateUserAsync(
                FirstnameTextbox.Text.Trim(),
                LastNameTextBox.Text.Trim(),
                PasswordTextbox.Password);

            if (isAuthenticated)
                {
                Window1 window1 = new Window1();
                window1.ShowDialog();
            }
            else
            {
                MessageBox.Show("ошибка, неверный пароль", "Ошибка входа",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            //Window1 window1 = new Window1();
            //window1.ShowDialog(); 

        }
        private async Task<bool> AuthenticateUserAsync(string firstName, string lastName, string password)
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"SELECT COUNT(*) FROM dbo.Employees 
                              WHERE FirstName = @FirstName 
                              AND LastName = @LastName 
                              AND Password = @Password";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@Password", password);
                int count = (int)await cmd.ExecuteScalarAsync();
                return count > 0;
            }
            catch (SqlException ex)
            {
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void OpenRegWindowButton_Click(object sender, RoutedEventArgs e)
        {
            Reg reg = new Reg();
            reg.ShowDialog();
            
        }
    }
}