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
using System.Data;
using Microsoft.Data.SqlClient;

namespace WpfApp9
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private const string ConnectionString = 
            "Server=localhost\\SQLEXPRESS;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";
        private const string Sql = "select * from dbo.Categories";
        
        // Хранение ID текущего пользователя
        private readonly int _currentEmployeeId;
        
        /// <summary>
        /// Конструктор с параметром ID пользователя
        /// </summary>
        /// <param name="employeeId">ID авторизованного пользователя</param>
        public Window1(int employeeId)
        {
            InitializeComponent();
            _currentEmployeeId = employeeId;
            Loaded += Window1_Loaded;
        }
        private async void Window1_Loaded(object sender, RoutedEventArgs e)
        {
            // Скрываем вкладку Employees для всех пользователей, кроме пользователя с ID = 1
            if (_currentEmployeeId != 1)
            {
                EmployeesTabItem.Visibility = Visibility.Collapsed;
            }
            
            await LoadAllTablesAsync();
        }
        private async Task LoadAllTablesAsync()
        {
            await LoadCategoriesTabAsync();
            await LoadEquipmentTabAsync(); 
            await LoadEmployeesTabAsync(); 
            await LoadWarehouseStockTabAsync();
        }
        private async Task LoadCategoriesTabAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT * FROM dbo.Categories";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                CategoriesTabGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
            }
            catch (Exception ex)
            {
            }
        }
        private async Task LoadEquipmentTabAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT * FROM dbo.Equipment";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                EquipmentTabGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex) 
            {
                MessageBox.Show($"SQL ошибка при загрузке продуктов: {ex.Message}", 
                    "SQL Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Ошибка при загрузке продуктов: {ex.Message}" +
                    $"", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task LoadEmployeesTabAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT EmployeeID,FirstName,LastName,Position,Email,Phone FROM dbo.Employees";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                EmployeesTabGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке продуктов: {ex.Message}",
                    "SQL Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке продуктов: {ex.Message}" +
                    $"", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task LoadWarehouseStockTabAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT * FROM dbo.WarehouseStock";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                WarehouseStockTabGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке продуктов: {ex.Message}",
                    "SQL Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке продуктов: {ex.Message}" +
                    $"", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public async Task RefreshTabAsync(string tabName)
        {
            switch (tabName)
            {
                case "Categories":
                    await LoadCategoriesTabAsync();
                    break;
                case "Equipment":
                    await LoadEquipmentTabAsync();
                    break;
            }
        }
        private void OpenAdminWindowButton_Click(object sender, RoutedEventArgs e)
        {
            AdminWindow adminwindow = new AdminWindow();
            adminwindow.ShowDialog();
        }
    }
}
