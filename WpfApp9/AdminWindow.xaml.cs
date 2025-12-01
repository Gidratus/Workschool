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

        // Хранение ID текущего пользователя
        private readonly int _currentEmployeeId;

        /// <summary>
        /// Конструктор с параметром ID пользователя
        /// </summary>
        /// <param name="employeeId">ID авторизованного пользователя</param>
        public AdminWindow(int employeeId)
        {
            InitializeComponent();
            _currentEmployeeId = employeeId;
            Loaded += AdminWindow_Loaded;
        }

        private async void AdminWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Скрываем вкладку Employees для всех пользователей, кроме пользователя с ID = 1
            if (_currentEmployeeId != 1)
            {
                EmployeesTabItem.Visibility = Visibility.Collapsed;
            }
            
            await LoadCategoriesAsync();
            await LoadCategoriesForEquipmentComboBoxAsync(); // Загружаем категории в ComboBox для Equipment
            await LoadEquipmentAsync(); // Загружаем данные оборудования
            
            // Загружаем данные Employees только если пользователь имеет доступ
            if (_currentEmployeeId == 1)
            {
                await LoadEmployeesAsync();
            }
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

        // ===============================================
        // Методы для работы с таблицей Employees
        // ===============================================

        /// <summary>
        /// Загрузка данных сотрудников из базы данных
        /// </summary>
        private async Task LoadEmployeesAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"SELECT EmployeeID, FirstName, LastName, Position, Email, Phone, Password 
                              FROM dbo.Employees ORDER BY EmployeeID";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                EmployeesDataGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке сотрудников: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке сотрудников: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обработчик выбора строки в таблице сотрудников
        /// </summary>
        private void EmployeesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EmployeesDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)EmployeesDataGrid.SelectedItem;
                EmployeeIdTextBox.Text = row["EmployeeID"].ToString();
                EmployeeFirstNameTextBox.Text = row["FirstName"].ToString();
                EmployeeLastNameTextBox.Text = row["LastName"].ToString();
                EmployeePositionTextBox.Text = row["Position"].ToString();
                EmployeeEmailTextBox.Text = row["Email"].ToString();
                EmployeePhoneTextBox.Text = row["Phone"].ToString();
                EmployeePasswordTextBox.Text = row["Password"].ToString();
            }
        }

        /// <summary>
        /// Добавление нового сотрудника
        /// </summary>
        private async void AddEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            // Валидация полей
            if (string.IsNullOrWhiteSpace(EmployeeFirstNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите имя сотрудника!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(EmployeeLastNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите фамилию сотрудника!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(EmployeePasswordTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите пароль!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"INSERT INTO dbo.Employees (FirstName, LastName, Position, Email, Phone, Password) 
                              VALUES (@FirstName, @LastName, @Position, @Email, @Phone, @Password)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", EmployeeFirstNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@LastName", EmployeeLastNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Position", 
                    string.IsNullOrWhiteSpace(EmployeePositionTextBox.Text) ? (object)DBNull.Value : EmployeePositionTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Email", 
                    string.IsNullOrWhiteSpace(EmployeeEmailTextBox.Text) ? (object)DBNull.Value : EmployeeEmailTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Phone", 
                    string.IsNullOrWhiteSpace(EmployeePhoneTextBox.Text) ? (object)DBNull.Value : EmployeePhoneTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Password", EmployeePasswordTextBox.Text.Trim());

                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Сотрудник успешно добавлен!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEmployeeFields();
                    await LoadEmployeesAsync();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при добавлении сотрудника: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении сотрудника: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обновление данных сотрудника
        /// </summary>
        private async void UpdateEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EmployeeIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите сотрудника для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(EmployeeFirstNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите имя сотрудника!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(EmployeeLastNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите фамилию сотрудника!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(EmployeePasswordTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите пароль!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"UPDATE dbo.Employees 
                              SET FirstName = @FirstName, LastName = @LastName, 
                                  Position = @Position, Email = @Email, Phone = @Phone, Password = @Password 
                              WHERE EmployeeID = @EmployeeID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FirstName", EmployeeFirstNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@LastName", EmployeeLastNameTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Position", 
                    string.IsNullOrWhiteSpace(EmployeePositionTextBox.Text) ? (object)DBNull.Value : EmployeePositionTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Email", 
                    string.IsNullOrWhiteSpace(EmployeeEmailTextBox.Text) ? (object)DBNull.Value : EmployeeEmailTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Phone", 
                    string.IsNullOrWhiteSpace(EmployeePhoneTextBox.Text) ? (object)DBNull.Value : EmployeePhoneTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Password", EmployeePasswordTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@EmployeeID", int.Parse(EmployeeIdTextBox.Text));

                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Данные сотрудника успешно обновлены!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEmployeeFields();
                    await LoadEmployeesAsync();
                }
                else
                {
                    MessageBox.Show("Сотрудник не найден!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при обновлении данных сотрудника: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных сотрудника: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Удаление сотрудника
        /// </summary>
        private async void DeleteEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EmployeeIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите сотрудника для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                $"Вы уверены, что хотите удалить сотрудника '{EmployeeFirstNameTextBox.Text} {EmployeeLastNameTextBox.Text}'?",
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
                string sql = "DELETE FROM dbo.Employees WHERE EmployeeID = @EmployeeID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", int.Parse(EmployeeIdTextBox.Text));
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Сотрудник успешно удален!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEmployeeFields();
                    await LoadEmployeesAsync();
                }
                else
                {
                    MessageBox.Show("Сотрудник не найден!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при удалении сотрудника: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении сотрудника: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Очистка полей формы сотрудника
        /// </summary>
        private void ClearEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            ClearEmployeeFields();
            EmployeesDataGrid.SelectedItem = null;
        }

        /// <summary>
        /// Обновление списка сотрудников
        /// </summary>
        private async void RefreshEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadEmployeesAsync();
            MessageBox.Show("Список сотрудников обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// Вспомогательный метод для очистки полей сотрудника
        /// </summary>
        private void ClearEmployeeFields()
        {
            EmployeeIdTextBox.Clear();
            EmployeeFirstNameTextBox.Clear();
            EmployeeLastNameTextBox.Clear();
            EmployeePositionTextBox.Clear();
            EmployeeEmailTextBox.Clear();
            EmployeePhoneTextBox.Clear();
            EmployeePasswordTextBox.Clear();
        }

        // ===============================================
        // Методы для работы с таблицей Equipment
        // ===============================================

        /// <summary>
        /// Загрузка категорий в ComboBox для выбора в форме Equipment
        /// </summary>
        private async Task LoadCategoriesForEquipmentComboBoxAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT CategoryID, CategoryName FROM dbo.Categories ORDER BY CategoryName";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                EquipmentCategoryComboBox.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке категорий для ComboBox: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Загрузка данных оборудования из базы данных
        /// </summary>
        private async Task LoadEquipmentAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // Загружаем все поля из таблицы Equipment
                string sql = @"SELECT EquipmentID, EquipmentName, CategoryID, Manufacturer, 
                              Model, SerialNumber, PurchaseDate, WarrantyUntil 
                              FROM dbo.Equipment ORDER BY EquipmentID";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                EquipmentDataGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке оборудования: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке оборудования: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обработчик выбора строки в таблице оборудования
        /// </summary>
        private void EquipmentDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EquipmentDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)EquipmentDataGrid.SelectedItem;
                // Заполняем все поля формы данными из выбранной строки
                EquipmentIdTextBox.Text = row["EquipmentID"].ToString();
                EquipmentNameTextBox.Text = row["EquipmentName"].ToString();
                // Устанавливаем выбранную категорию в ComboBox
                EquipmentCategoryComboBox.SelectedValue = row["CategoryID"] != DBNull.Value ? row["CategoryID"] : null;
                EquipmentManufacturerTextBox.Text = row["Manufacturer"] != DBNull.Value ? row["Manufacturer"].ToString() : "";
                EquipmentModelTextBox.Text = row["Model"] != DBNull.Value ? row["Model"].ToString() : "";
                EquipmentSerialNumberTextBox.Text = row["SerialNumber"] != DBNull.Value ? row["SerialNumber"].ToString() : "";
                // Устанавливаем даты в DatePicker
                EquipmentPurchaseDatePicker.SelectedDate = row["PurchaseDate"] != DBNull.Value 
                    ? (DateTime?)row["PurchaseDate"] : null;
                EquipmentWarrantyUntilPicker.SelectedDate = row["WarrantyUntil"] != DBNull.Value 
                    ? (DateTime?)row["WarrantyUntil"] : null;
            }
        }

        /// <summary>
        /// Добавление нового оборудования
        /// </summary>
        private async void AddEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            // Валидация - проверка, что название введено
            if (string.IsNullOrWhiteSpace(EquipmentNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите название оборудования!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // INSERT со всеми полями
                string sql = @"INSERT INTO dbo.Equipment 
                              (EquipmentName, CategoryID, Manufacturer, Model, SerialNumber, PurchaseDate, WarrantyUntil) 
                              VALUES (@EquipmentName, @CategoryID, @Manufacturer, @Model, @SerialNumber, @PurchaseDate, @WarrantyUntil)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentName", EquipmentNameTextBox.Text.Trim());
                // CategoryID - берём из ComboBox
                cmd.Parameters.AddWithValue("@CategoryID", 
                    EquipmentCategoryComboBox.SelectedValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Manufacturer", 
                    string.IsNullOrWhiteSpace(EquipmentManufacturerTextBox.Text) ? (object)DBNull.Value : EquipmentManufacturerTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Model", 
                    string.IsNullOrWhiteSpace(EquipmentModelTextBox.Text) ? (object)DBNull.Value : EquipmentModelTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@SerialNumber", 
                    string.IsNullOrWhiteSpace(EquipmentSerialNumberTextBox.Text) ? (object)DBNull.Value : EquipmentSerialNumberTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@PurchaseDate", 
                    EquipmentPurchaseDatePicker.SelectedDate.HasValue ? (object)EquipmentPurchaseDatePicker.SelectedDate.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@WarrantyUntil", 
                    EquipmentWarrantyUntilPicker.SelectedDate.HasValue ? (object)EquipmentWarrantyUntilPicker.SelectedDate.Value : DBNull.Value);
                
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Оборудование успешно добавлено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEquipmentFields();
                    await LoadEquipmentAsync();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при добавлении оборудования: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении оборудования: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обновление данных оборудования
        /// </summary>
        private async void UpdateEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что выбрана запись для обновления
            if (string.IsNullOrWhiteSpace(EquipmentIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите оборудование для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            // Проверка, что название введено
            if (string.IsNullOrWhiteSpace(EquipmentNameTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, введите название оборудования!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // UPDATE со всеми полями
                string sql = @"UPDATE dbo.Equipment SET 
                              EquipmentName = @EquipmentName, 
                              CategoryID = @CategoryID, 
                              Manufacturer = @Manufacturer, 
                              Model = @Model, 
                              SerialNumber = @SerialNumber, 
                              PurchaseDate = @PurchaseDate, 
                              WarrantyUntil = @WarrantyUntil 
                              WHERE EquipmentID = @EquipmentID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentName", EquipmentNameTextBox.Text.Trim());
                // CategoryID - берём из ComboBox
                cmd.Parameters.AddWithValue("@CategoryID", 
                    EquipmentCategoryComboBox.SelectedValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Manufacturer", 
                    string.IsNullOrWhiteSpace(EquipmentManufacturerTextBox.Text) ? (object)DBNull.Value : EquipmentManufacturerTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@Model", 
                    string.IsNullOrWhiteSpace(EquipmentModelTextBox.Text) ? (object)DBNull.Value : EquipmentModelTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@SerialNumber", 
                    string.IsNullOrWhiteSpace(EquipmentSerialNumberTextBox.Text) ? (object)DBNull.Value : EquipmentSerialNumberTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@PurchaseDate", 
                    EquipmentPurchaseDatePicker.SelectedDate.HasValue ? (object)EquipmentPurchaseDatePicker.SelectedDate.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@WarrantyUntil", 
                    EquipmentWarrantyUntilPicker.SelectedDate.HasValue ? (object)EquipmentWarrantyUntilPicker.SelectedDate.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@EquipmentID", int.Parse(EquipmentIdTextBox.Text));
                
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Оборудование успешно обновлено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEquipmentFields();
                    await LoadEquipmentAsync();
                }
                else
                {
                    MessageBox.Show("Оборудование не найдено!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при обновлении оборудования: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении оборудования: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Удаление оборудования
        /// </summary>
        private async void DeleteEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что выбрана запись для удаления
            if (string.IsNullOrWhiteSpace(EquipmentIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите оборудование для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            // Подтверждение удаления
            var result = MessageBox.Show(
                $"Вы уверены, что хотите удалить оборудование '{EquipmentNameTextBox.Text}'?",
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
                string sql = "DELETE FROM dbo.Equipment WHERE EquipmentID = @EquipmentID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentID", int.Parse(EquipmentIdTextBox.Text));
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Оборудование успешно удалено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearEquipmentFields();
                    await LoadEquipmentAsync();
                }
                else
                {
                    MessageBox.Show("Оборудование не найдено!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при удалении оборудования: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении оборудования: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Очистка полей формы оборудования
        /// </summary>
        private void ClearEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            ClearEquipmentFields();
            EquipmentDataGrid.SelectedItem = null;
        }

        /// <summary>
        /// Обновление списка оборудования
        /// </summary>
        private async void RefreshEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadEquipmentAsync();
            MessageBox.Show("Список оборудования обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// Вспомогательный метод для очистки полей оборудования
        /// </summary>
        private void ClearEquipmentFields()
        {
            EquipmentIdTextBox.Clear();
            EquipmentNameTextBox.Clear();
            EquipmentCategoryComboBox.SelectedIndex = -1; // Сбрасываем выбор в ComboBox
            EquipmentManufacturerTextBox.Clear();
            EquipmentModelTextBox.Clear();
            EquipmentSerialNumberTextBox.Clear();
            EquipmentPurchaseDatePicker.SelectedDate = null;
            EquipmentWarrantyUntilPicker.SelectedDate = null;
        }
    }
}
