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
    public class UserData
    {
        public int EmployeeId { get; set; }
        public string Position { get; set; }
    }
    public partial class MainWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;
        private const string Sql = "select * from dbo.Categories";
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
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
            UserData userData = await AuthenticateUserAsync(
                FirstnameTextbox.Text.Trim(),
                LastNameTextBox.Text.Trim(),
                PasswordTextbox.Password);
            if (userData != null)
            {
                Window1 window1 = new Window1(userData);
                window1.ShowDialog();
            }
            else
            {
                MessageBox.Show("ошибка, неверный пароль", "Ошибка входа",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task<UserData> AuthenticateUserAsync(string firstName, string lastName, string password)
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"SELECT EmployeeID, Position FROM dbo.Employees 
                              WHERE FirstName = @FirstName 
                              AND LastName = @LastName 
                              AND Password = @Password";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@Password", password);
                await using var reader = await cmd.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new UserData
                    {
                        EmployeeId = reader.GetInt32(0),
                        Position = reader.IsDBNull(1) ? "" : reader.GetString(1)
                    };
                }
                return null;
            }
            catch (SqlException ex)
            {
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private void OpenRegWindowButton_Click(object sender, RoutedEventArgs e)
        {
            Reg reg = new Reg();
            reg.ShowDialog();
        }
    }
}
