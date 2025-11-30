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
            
            // Получаем ID пользователя при авторизации
            int? employeeId = await AuthenticateUserAsync(
                FirstnameTextbox.Text.Trim(),
                LastNameTextBox.Text.Trim(),
                PasswordTextbox.Password);

            if (employeeId.HasValue)
            {
                // Передаем ID пользователя в Window1
                Window1 window1 = new Window1(employeeId.Value);
                window1.ShowDialog();
            }
            else
            {
                MessageBox.Show("ошибка, неверный пароль", "Ошибка входа",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// Аутентификация пользователя и получение его ID
        /// </summary>
        /// <returns>EmployeeID если авторизация успешна, null если нет</returns>
        private async Task<int?> AuthenticateUserAsync(string firstName, string lastName, string password)
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                
                // Получаем EmployeeID вместо COUNT(*)
                string sql = @"SELECT EmployeeID FROM dbo.Employees 
                              WHERE FirstName = @FirstName 
                              AND LastName = @LastName 
                              AND Password = @Password";
                              
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@Password", password);
                
                var result = await cmd.ExecuteScalarAsync();
                
                // Если пользователь найден, возвращаем его ID
                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt32(result);
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