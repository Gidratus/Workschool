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
        private const string ConnectionString = App.ConnectionString;
        private const string Sql = "select * from dbo.Categories";
        
        // Хранение данных текущего пользователя
        private readonly UserData _currentUser;
        
        /// <summary>
        /// Конструктор с параметром данных пользователя
        /// </summary>
        /// <param name="userData">Данные авторизованного пользователя (ID и должность)</param>
        public Window1(UserData userData)
        {
            InitializeComponent();
            _currentUser = userData;
            Loaded += Window1_Loaded;
        }
        private async void Window1_Loaded(object sender, RoutedEventArgs e)
        {
            // Скрываем вкладку Employees для всех пользователей, кроме пользователя с ID = 1
            if (_currentUser.EmployeeId != 1)
            {
                EmployeesTabItem.Visibility = Visibility.Collapsed;
            }
            
            // Скрываем кнопку "Админ доступ" и "Подтверждение работ" для стажеров
            // Эти функции доступны только для Менеджеров и Администраторов
            if (_currentUser.Position != null && _currentUser.Position.Equals("Стажер", StringComparison.OrdinalIgnoreCase))
            {
                AdminAccessButton.Visibility = Visibility.Collapsed;
                WorkConfirmationButton.Visibility = Visibility.Collapsed;
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
            // Передаем ID текущего пользователя в AdminWindow
            AdminWindow adminwindow = new AdminWindow(_currentUser.EmployeeId);
            adminwindow.ShowDialog();
        }


        private void OpenProfileButton_Click(object sender, RoutedEventArgs e)
        {
            ProfileWindow profileWindow = new ProfileWindow(_currentUser.EmployeeId);
            profileWindow.Owner = this;
            profileWindow.ShowDialog();
        }

        /// <summary>
        /// Открытие окна просмотра записей движений оборудования для текущего пользователя
        /// </summary>
        private void OpenMyMovementsButton_Click(object sender, RoutedEventArgs e)
        {
            MyMovementsWindow myMovementsWindow = new MyMovementsWindow(_currentUser.EmployeeId);
            myMovementsWindow.Owner = this;
            myMovementsWindow.ShowDialog();
        }

        /// <summary>
        /// Открытие окна подтверждения работ для Менеджеров и Администраторов
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Создаем и открываем окно подтверждения работ
                // Передаем данные текущего пользователя для проверки прав доступа
                WorkConfirmationWindow confirmationWindow = new WorkConfirmationWindow(_currentUser);
                confirmationWindow.Owner = this;
                confirmationWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии окна: {ex.Message}\n\nТип: {ex.GetType().Name}\n\nПодробности: {ex.StackTrace}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
