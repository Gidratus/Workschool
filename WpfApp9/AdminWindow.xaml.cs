using System;
using System.Collections.Generic;
using System.Data;
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

    public partial class AdminWindow : Window
    {
        private const string ConnectionString =
            "Server=localhost\\SQLEXPRESS;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";

        public AdminWindow()
        {
            InitializeComponent();
            Loaded += AdminWindow_Loaded;
        }

        private async void AdminWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadCategoriesAsync();
        }

        private async Task LoadCategoriesAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT CategoryID, CategoryName FROM dbo.Categories ORDER BY CategoryID";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                CategoriesDataGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке категорий: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке категорий: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CategoriesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoriesDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)CategoriesDataGrid.SelectedItem;   
                CategoryIdTextBox.Text = row["CategoryID"].ToString();
                CategoryNameTextBox.Text = row["CategoryName"].ToString();
            }
        }

        private async void AddButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CategoryNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите название категории!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "INSERT INTO dbo.Categories (CategoryName) VALUES (@CategoryName)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@CategoryName", CategoryNameTextBox.Text.Trim());
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Категория успешно добавлена!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearFields();
                    await LoadCategoriesAsync();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при добавлении категории: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении категории: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CategoryIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите категорию для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(CategoryNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите название категории!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "UPDATE dbo.Categories SET CategoryName = @CategoryName WHERE CategoryID = @CategoryID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@CategoryName", CategoryNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@CategoryID", int.Parse(CategoryIdTextBox.Text));
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Категория успешно обновлена!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearFields();
                    await LoadCategoriesAsync();
                }
                else
                {
                    MessageBox.Show("Категория не найдена!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при обновлении категории: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении категории: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CategoryIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите категорию для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var result = MessageBox.Show(
                $"Вы уверены, что хотите удалить категорию '{CategoryNameTextBox.Text}'?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "DELETE FROM dbo.Categories WHERE CategoryID = @CategoryID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@CategoryID", int.Parse(CategoryIdTextBox.Text));
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Категория успешно удалена!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearFields();
                    await LoadCategoriesAsync();
                }
                else
                {
                    MessageBox.Show("Категория не найдена!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                // if (ex.Number == 547)
                // {
                //     MessageBox.Show(
                //         "Невозможно удалить категорию, так как она используется в других таблицах!",
                //         "Ошибка удаления", MessageBoxButton.OK, MessageBoxImage.Error);
                // }
                // else
                // {
                //     MessageBox.Show($"SQL ошибка при удалении категории: {ex.Message}",
                //         "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
                // }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении категории: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearFields();
            CategoriesDataGrid.SelectedItem = null;
        }
        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadCategoriesAsync();
            MessageBox.Show("Список категорий обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void ClearFields()
        {
            CategoryIdTextBox.Clear();
            CategoryNameTextBox.Clear();
        }
    }
}
