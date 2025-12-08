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
using Microsoft.Win32;
// Пространства имён для работы с Word документами (OpenXML)
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WpfApp9
{

    public partial class AdminWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;

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
            
            // Загружаем данные для вкладки движений оборудования
            await LoadEquipmentForMovementComboBoxAsync();
            await LoadSuppliersForMovementComboBoxAsync();
            await LoadEmployeesForMovementComboBoxAsync();
            await LoadMovementAsync();
            
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

        /// <summary>
        /// Обработчик кнопки "Сохранить Word" - создаёт Word документ с заголовком "Отчет"
        /// </summary>
        private void SaveCategoriesWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Открываем диалог сохранения файла
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "Отчет_Категории"
                };

                // Если пользователь выбрал файл и нажал "Сохранить"
                if (saveFileDialog.ShowDialog() == true)
                {
                    // Создаём Word документ с заголовком "Отчет"
                    CreateWordDocument(saveFileDialog.FileName);
                    
                    MessageBox.Show("Документ Word успешно сохранен!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении документа: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Создаёт Word документ с заголовком "Отчет"
        /// </summary>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        private void CreateWordDocument(string filePath)
        {
            // Создаём новый Word документ
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Добавляем основную часть документа
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Создаём параграф с заголовком "Отчет"
                Paragraph titleParagraph = new Paragraph();
                
                // Настройки параграфа - выравнивание по центру
                ParagraphProperties paragraphProperties = new ParagraphProperties();
                Justification justification = new Justification() { Val = JustificationValues.Center };
                paragraphProperties.Append(justification);
                titleParagraph.Append(paragraphProperties);

                // Создаём текст заголовка
                Run run = new Run();
                
                // Настройки текста - жирный шрифт, размер 28pt (28 * 2 = 56 в OpenXML)
                RunProperties runProperties = new RunProperties();
                Bold bold = new Bold();
                FontSize fontSize = new FontSize() { Val = "56" };
                runProperties.Append(bold);
                runProperties.Append(fontSize);
                run.Append(runProperties);
                
                // Добавляем текст "Отчет"
                Text text = new Text("Отчет");
                run.Append(text);
                
                titleParagraph.Append(run);
                body.Append(titleParagraph);

                // Сохраняем документ
                mainPart.Document.Save();
            }
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

        // ===============================================
        // Методы для работы с таблицей EquipmentMovement
        // ===============================================

        /// <summary>
        /// Загрузка оборудования в ComboBox для формы движений
        /// </summary>
        private async Task LoadEquipmentForMovementComboBoxAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT EquipmentID, EquipmentName FROM dbo.Equipment ORDER BY EquipmentName";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                MovementEquipmentComboBox.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке оборудования для ComboBox: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Загрузка поставщиков в ComboBox для формы движений
        /// </summary>
        private async Task LoadSuppliersForMovementComboBoxAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = "SELECT SupplierID, SupplierName FROM dbo.Suppliers ORDER BY SupplierName";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                MovementSupplierComboBox.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке поставщиков для ComboBox: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Загрузка сотрудников в ComboBox для формы движений
        /// </summary>
        private async Task LoadEmployeesForMovementComboBoxAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // Объединяем имя и фамилию для отображения
                string sql = @"SELECT EmployeeID, FirstName + ' ' + LastName AS FullName 
                              FROM dbo.Employees ORDER BY LastName, FirstName";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                MovementEmployeeComboBox.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке сотрудников для ComboBox: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Загрузка данных движений оборудования из базы данных
        /// </summary>
        private async Task LoadMovementAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // Загружаем движения с JOIN для получения имен оборудования, поставщика и сотрудника
                string sql = @"SELECT m.MovementID, m.EquipmentID, e.EquipmentName, 
                              m.MovementDate, m.Quantity, m.MovementType, 
                              m.SupplierID, s.SupplierName,
                              m.EmployeeID, emp.FirstName + ' ' + emp.LastName AS EmployeeName,
                              m.Notes
                              FROM dbo.EquipmentMovement m
                              LEFT JOIN dbo.Equipment e ON m.EquipmentID = e.EquipmentID
                              LEFT JOIN dbo.Suppliers s ON m.SupplierID = s.SupplierID
                              LEFT JOIN dbo.Employees emp ON m.EmployeeID = emp.EmployeeID
                              ORDER BY m.MovementID DESC";
                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                MovementDataGrid.ItemsSource = dt.DefaultView;
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке движений: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке движений: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обработчик выбора строки в таблице движений
        /// </summary>
        private void MovementDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MovementDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)MovementDataGrid.SelectedItem;
                // Заполняем все поля формы данными из выбранной строки
                MovementIdTextBox.Text = row["MovementID"].ToString();
                // Устанавливаем оборудование в ComboBox
                MovementEquipmentComboBox.SelectedValue = row["EquipmentID"] != DBNull.Value ? row["EquipmentID"] : null;
                // Устанавливаем дату
                MovementDatePicker.SelectedDate = row["MovementDate"] != DBNull.Value 
                    ? (DateTime?)row["MovementDate"] : null;
                // Количество
                MovementQuantityTextBox.Text = row["Quantity"] != DBNull.Value ? row["Quantity"].ToString() : "";
                // Устанавливаем тип движения в ComboBox
                string movementType = row["MovementType"]?.ToString() ?? "";
                foreach (ComboBoxItem item in MovementTypeComboBox.Items)
                {
                    if (item.Content?.ToString() == movementType)
                    {
                        MovementTypeComboBox.SelectedItem = item;
                        break;
                    }
                }
                // Поставщик
                MovementSupplierComboBox.SelectedValue = row["SupplierID"] != DBNull.Value ? row["SupplierID"] : null;
                // Сотрудник
                MovementEmployeeComboBox.SelectedValue = row["EmployeeID"] != DBNull.Value ? row["EmployeeID"] : null;
                // Примечания
                MovementNotesTextBox.Text = row["Notes"] != DBNull.Value ? row["Notes"].ToString() : "";
            }
        }

        /// <summary>
        /// Добавление нового движения оборудования
        /// </summary>
        private async void AddMovementButton_Click(object sender, RoutedEventArgs e)
        {
            // Валидация - проверка обязательных полей
            if (MovementEquipmentComboBox.SelectedValue == null)
            {
                MessageBox.Show("Пожалуйста, выберите оборудование!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!MovementDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Пожалуйста, выберите дату движения!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (MovementTypeComboBox.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите тип движения!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Проверка количества на число
            int? quantity = null;
            if (!string.IsNullOrWhiteSpace(MovementQuantityTextBox.Text))
            {
                if (!int.TryParse(MovementQuantityTextBox.Text, out int parsedQuantity))
                {
                    MessageBox.Show("Количество должно быть целым числом!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                quantity = parsedQuantity;
            }

            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // INSERT со всеми полями
                string sql = @"INSERT INTO dbo.EquipmentMovement 
                              (EquipmentID, MovementDate, Quantity, MovementType, SupplierID, Notes, EmployeeID) 
                              VALUES (@EquipmentID, @MovementDate, @Quantity, @MovementType, @SupplierID, @Notes, @EmployeeID)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentID", MovementEquipmentComboBox.SelectedValue);
                cmd.Parameters.AddWithValue("@MovementDate", MovementDatePicker.SelectedDate.Value);
                cmd.Parameters.AddWithValue("@Quantity", quantity.HasValue ? (object)quantity.Value : DBNull.Value);
                // Получаем текст выбранного типа движения
                string movementType = ((ComboBoxItem)MovementTypeComboBox.SelectedItem).Content.ToString();
                cmd.Parameters.AddWithValue("@MovementType", movementType);
                cmd.Parameters.AddWithValue("@SupplierID", 
                    MovementSupplierComboBox.SelectedValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Notes", 
                    string.IsNullOrWhiteSpace(MovementNotesTextBox.Text) ? (object)DBNull.Value : MovementNotesTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@EmployeeID", 
                    MovementEmployeeComboBox.SelectedValue ?? DBNull.Value);
                
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Движение оборудования успешно добавлено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearMovementFields();
                    await LoadMovementAsync();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при добавлении движения: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении движения: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обновление данных движения оборудования
        /// </summary>
        private async void UpdateMovementButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что выбрана запись для обновления
            if (string.IsNullOrWhiteSpace(MovementIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите движение для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            // Валидация обязательных полей
            if (MovementEquipmentComboBox.SelectedValue == null)
            {
                MessageBox.Show("Пожалуйста, выберите оборудование!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!MovementDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Пожалуйста, выберите дату движения!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (MovementTypeComboBox.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите тип движения!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Проверка количества на число
            int? quantity = null;
            if (!string.IsNullOrWhiteSpace(MovementQuantityTextBox.Text))
            {
                if (!int.TryParse(MovementQuantityTextBox.Text, out int parsedQuantity))
                {
                    MessageBox.Show("Количество должно быть целым числом!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                quantity = parsedQuantity;
            }

            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                // UPDATE со всеми полями
                string sql = @"UPDATE dbo.EquipmentMovement SET 
                              EquipmentID = @EquipmentID, 
                              MovementDate = @MovementDate, 
                              Quantity = @Quantity, 
                              MovementType = @MovementType, 
                              SupplierID = @SupplierID, 
                              Notes = @Notes,
                              EmployeeID = @EmployeeID
                              WHERE MovementID = @MovementID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentID", MovementEquipmentComboBox.SelectedValue);
                cmd.Parameters.AddWithValue("@MovementDate", MovementDatePicker.SelectedDate.Value);
                cmd.Parameters.AddWithValue("@Quantity", quantity.HasValue ? (object)quantity.Value : DBNull.Value);
                // Получаем текст выбранного типа движения
                string movementType = ((ComboBoxItem)MovementTypeComboBox.SelectedItem).Content.ToString();
                cmd.Parameters.AddWithValue("@MovementType", movementType);
                cmd.Parameters.AddWithValue("@SupplierID", 
                    MovementSupplierComboBox.SelectedValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Notes", 
                    string.IsNullOrWhiteSpace(MovementNotesTextBox.Text) ? (object)DBNull.Value : MovementNotesTextBox.Text.Trim());
                cmd.Parameters.AddWithValue("@EmployeeID", 
                    MovementEmployeeComboBox.SelectedValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@MovementID", int.Parse(MovementIdTextBox.Text));
                
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Движение оборудования успешно обновлено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearMovementFields();
                    await LoadMovementAsync();
                }
                else
                {
                    MessageBox.Show("Движение не найдено!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при обновлении движения: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении движения: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Удаление движения оборудования
        /// </summary>
        private async void DeleteMovementButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что выбрана запись для удаления
            if (string.IsNullOrWhiteSpace(MovementIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите движение для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            // Подтверждение удаления
            var result = MessageBox.Show(
                "Вы уверены, что хотите удалить это движение оборудования?",
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
                string sql = "DELETE FROM dbo.EquipmentMovement WHERE MovementID = @MovementID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@MovementID", int.Parse(MovementIdTextBox.Text));
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    MessageBox.Show("Движение оборудования успешно удалено!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    ClearMovementFields();
                    await LoadMovementAsync();
                }
                else
                {
                    MessageBox.Show("Движение не найдено!",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при удалении движения: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении движения: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Очистка полей формы движений
        /// </summary>
        private void ClearMovementButton_Click(object sender, RoutedEventArgs e)
        {
            ClearMovementFields();
            MovementDataGrid.SelectedItem = null;
        }

        /// <summary>
        /// Обновление списка движений
        /// </summary>
        private async void RefreshMovementButton_Click(object sender, RoutedEventArgs e)
        {
            // Обновляем также ComboBox-ы на случай добавления нового оборудования/поставщиков/сотрудников
            await LoadEquipmentForMovementComboBoxAsync();
            await LoadSuppliersForMovementComboBoxAsync();
            await LoadEmployeesForMovementComboBoxAsync();
            await LoadMovementAsync();
            MessageBox.Show("Список движений оборудования обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// Вспомогательный метод для очистки полей движений
        /// </summary>
        private void ClearMovementFields()
        {
            MovementIdTextBox.Clear();
            MovementEquipmentComboBox.SelectedIndex = -1;
            MovementDatePicker.SelectedDate = null;
            MovementQuantityTextBox.Clear();
            MovementTypeComboBox.SelectedIndex = -1;
            MovementSupplierComboBox.SelectedIndex = -1;
            MovementEmployeeComboBox.SelectedIndex = -1;
            MovementNotesTextBox.Clear();
        }
    }
}
