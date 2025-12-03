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
using Microsoft.Data.SqlClient;


namespace WpfApp9
{
    /// <summary>
    /// Логика взаимодействия для Reg.xaml
    /// </summary>
    public partial class Reg : Window
    {
        private const string ConnectionString = App.ConnectionString;

        public Reg()
        {
            InitializeComponent();
        }
        private async void RegisterButton_Click(object sender, RoutedEventArgs e)
        {
            //if (!ValidateInput())
            //{
            //    return;
            //}
            if (PasswordBox.Password != SecondPasswordBox.Password)
            {
                MessageBox.Show("Пароли не совпадают!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            await RegisterUserAsync();
        }
        private void ClearFields()
        {
            FirstNameTextBox.Clear();
            LastNameTextBox.Clear();
            EmailTextBox.Clear();
            PhoneTextBox.Clear();
            PasswordBox.Clear();
            SecondPasswordBox.Clear();
        }
        private async Task RegisterUserAsync()
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"INSERT INTO dbo.Employees (FirstName, LastName, Email, Phone, Password, Position) 
                              VALUES (@FirstName, @LastName, @Email, @Phone, @Password, @Position)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", FirstNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@LastName", LastNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Email", EmailTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Phone", PhoneTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Password", PasswordBox.Password);
                cmd.Parameters.AddWithValue("@Position", "Стажер");
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Пользователь успешно зарегистрирован!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearFields();
                }
                else
                {
                    MessageBox.Show("Не удалось зарегистрировать пользователя!", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при регистрации: {ex.Message}", "SQL Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при регистрации: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
