using System;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;
namespace WpfApp9
{
    public partial class ProfileWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;
        private readonly int _employeeId;
        public ProfileWindow(int employeeId)
        {
            InitializeComponent();
            _employeeId = employeeId;
            Loaded += ProfileWindow_Loaded;
        }
        private async void ProfileWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadUserDataAsync();
        }
        private async Task LoadUserDataAsync()
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"SELECT FirstName, LastName, Email 
                              FROM dbo.Employees 
                              WHERE EmployeeID = @EmployeeID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", _employeeId);
                await using var reader = await cmd.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    FirstNameTextBox.Text = reader.IsDBNull(0) ? "" : reader.GetString(0);
                    LastNameTextBox.Text = reader.IsDBNull(1) ? "" : reader.GetString(1);
                    EmailTextBox.Text = reader.IsDBNull(2) ? "" : reader.GetString(2);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке данных профиля: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных профиля: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(FirstNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите имя!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(LastNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите фамилию!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql;
                SqlCommand cmd;
                if (!string.IsNullOrWhiteSpace(PasswordBox.Password))
                {
                    sql = @"UPDATE dbo.Employees 
                           SET FirstName = @FirstName, 
                               LastName = @LastName, 
                               Email = @Email,
                               Password = @Password
                           WHERE EmployeeID = @EmployeeID";
                    cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@Password", PasswordBox.Password);
                }
                else
                {
                    sql = @"UPDATE dbo.Employees 
                           SET FirstName = @FirstName, 
                               LastName = @LastName, 
                               Email = @Email
                           WHERE EmployeeID = @EmployeeID";
                    cmd = new SqlCommand(sql, conn);
                }
                cmd.Parameters.AddWithValue("@FirstName", FirstNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@LastName", LastNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Email",
                    string.IsNullOrWhiteSpace(EmailTextBox.Text) ? (object)DBNull.Value : EmailTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@EmployeeID", _employeeId);
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Данные профиля успешно сохранены!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    DialogResult = true;
                    Close();
                }
                else
                {
                    MessageBox.Show("Не удалось сохранить данные профиля.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при сохранении профиля: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении профиля: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
