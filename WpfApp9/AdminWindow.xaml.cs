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
// Используем алиас для избежания конфликта имён с System.Windows.Documents
using WordProcessing = DocumentFormat.OpenXml.Wordprocessing;

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
        /// Обработчик кнопки "Сохранить Word" - создаёт Word документ с таблицей категорий
        /// </summary>
        private async void SaveCategoriesWordButton_Click(object sender, RoutedEventArgs e)
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
                    // Получаем данные категорий из БД
                    DataTable categoriesData = await GetCategoriesDataAsync();
                    
                    // Создаём Word документ с заголовком "Отчет" и таблицей категорий
                    CreateWordDocumentWithCategories(saveFileDialog.FileName, categoriesData);
                    
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
        /// Получает данные категорий из базы данных
        /// </summary>
        /// <returns>DataTable с категориями</returns>
        private async Task<DataTable> GetCategoriesDataAsync()
        {
            var dt = new DataTable();
            await using var conn = new SqlConnection(ConnectionString);
            await conn.OpenAsync();
            string sql = "SELECT CategoryID, CategoryName FROM dbo.Categories ORDER BY CategoryID";
            await using var cmd = new SqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
            return dt;
        }

        /// <summary>
        /// Создаёт Word документ с заголовком "Отчет" и таблицей категорий
        /// </summary>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        /// <param name="categoriesData">Данные категорий для таблицы</param>
        private void CreateWordDocumentWithCategories(string filePath, DataTable categoriesData)
        {
            // Создаём новый Word документ
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Добавляем основную часть документа
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());

                // === ЗАГОЛОВОК "Отчет" ===
                WordProcessing.Paragraph titleParagraph = new WordProcessing.Paragraph();
                
                // Настройки параграфа - выравнивание по центру
                WordProcessing.ParagraphProperties titleParagraphProps = new WordProcessing.ParagraphProperties();
                titleParagraphProps.Append(new WordProcessing.Justification() { Val = WordProcessing.JustificationValues.Center });
                // Отступ после заголовка
                titleParagraphProps.Append(new WordProcessing.SpacingBetweenLines() { After = "400" });
                titleParagraph.Append(titleParagraphProps);

                // Текст заголовка - жирный, 28pt
                WordProcessing.Run titleRun = new WordProcessing.Run();
                WordProcessing.RunProperties titleRunProps = new WordProcessing.RunProperties();
                titleRunProps.Append(new WordProcessing.Bold());
                titleRunProps.Append(new WordProcessing.FontSize() { Val = "56" }); // 28pt * 2
                titleRun.Append(titleRunProps);
                titleRun.Append(new WordProcessing.Text("Отчет"));
                
                titleParagraph.Append(titleRun);
                body.Append(titleParagraph);

                // === ПОДЗАГОЛОВОК "Категории" ===
                WordProcessing.Paragraph subtitleParagraph = new WordProcessing.Paragraph();
                WordProcessing.ParagraphProperties subtitleProps = new WordProcessing.ParagraphProperties();
                subtitleProps.Append(new WordProcessing.SpacingBetweenLines() { After = "200" });
                subtitleParagraph.Append(subtitleProps);
                
                WordProcessing.Run subtitleRun = new WordProcessing.Run();
                WordProcessing.RunProperties subtitleRunProps = new WordProcessing.RunProperties();
                subtitleRunProps.Append(new WordProcessing.Bold());
                subtitleRunProps.Append(new WordProcessing.FontSize() { Val = "32" }); // 16pt * 2
                subtitleRun.Append(subtitleRunProps);
                subtitleRun.Append(new WordProcessing.Text("Категории"));
                
                subtitleParagraph.Append(subtitleRun);
                body.Append(subtitleParagraph);

                // === ТАБЛИЦА С КАТЕГОРИЯМИ ===
                WordProcessing.Table table = new WordProcessing.Table();

                // Настройки таблицы - границы
                WordProcessing.TableProperties tableProps = new WordProcessing.TableProperties();
                WordProcessing.TableBorders tableBorders = new WordProcessing.TableBorders(
                    new WordProcessing.TopBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                    new WordProcessing.BottomBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                    new WordProcessing.LeftBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                    new WordProcessing.RightBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                    new WordProcessing.InsideHorizontalBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                    new WordProcessing.InsideVerticalBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 }
                );
                tableProps.Append(tableBorders);
                // Ширина таблицы - 100%
                tableProps.Append(new WordProcessing.TableWidth() { Width = "5000", Type = WordProcessing.TableWidthUnitValues.Pct });
                table.Append(tableProps);

                // === ЗАГОЛОВОК ТАБЛИЦЫ ===
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Название категории", true));
                table.Append(headerRow);

                // === СТРОКИ С ДАННЫМИ ===
                foreach (DataRow row in categoriesData.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["CategoryID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["CategoryName"].ToString() ?? "", false));
                    table.Append(dataRow);
                }

                body.Append(table);

                // Сохраняем документ
                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// Создаёт ячейку таблицы Word с текстом
        /// </summary>
        /// <param name="text">Текст ячейки</param>
        /// <param name="isHeader">Является ли ячейка заголовком (жирный текст)</param>
        /// <returns>Ячейка таблицы</returns>
        private WordProcessing.TableCell CreateTableCell(string text, bool isHeader)
        {
            WordProcessing.TableCell cell = new WordProcessing.TableCell();
            
            // Настройки ячейки - отступы внутри
            WordProcessing.TableCellProperties cellProps = new WordProcessing.TableCellProperties();
            cellProps.Append(new WordProcessing.TableCellMargin(
                new WordProcessing.TopMargin() { Width = "50", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.BottomMargin() { Width = "50", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.LeftMargin() { Width = "100", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.RightMargin() { Width = "100", Type = WordProcessing.TableWidthUnitValues.Dxa }
            ));
            
            // Для заголовка - серый фон
            if (isHeader)
            {
                cellProps.Append(new WordProcessing.Shading() 
                { 
                    Val = WordProcessing.ShadingPatternValues.Clear, 
                    Fill = "DDDDDD" // Светло-серый фон
                });
            }
            cell.Append(cellProps);

            // Параграф с текстом
            WordProcessing.Paragraph paragraph = new WordProcessing.Paragraph();
            WordProcessing.Run run = new WordProcessing.Run();
            
            // Настройки текста
            WordProcessing.RunProperties runProps = new WordProcessing.RunProperties();
            runProps.Append(new WordProcessing.FontSize() { Val = "24" }); // 12pt
            if (isHeader)
            {
                runProps.Append(new WordProcessing.Bold());
            }
            run.Append(runProps);
            run.Append(new WordProcessing.Text(text));
            
            paragraph.Append(run);
            cell.Append(paragraph);
            
            return cell;
        }

        // ===============================================
        // Экспорт в Word для Сотрудников
        // ===============================================

        /// <summary>
        /// Обработчик кнопки "Сохранить Word" для сотрудников
        /// </summary>
        private async void SaveEmployeesWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "Отчет_Сотрудники"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Получаем данные сотрудников из БД
                    DataTable employeesData = await GetEmployeesDataAsync();
                    
                    // Создаём Word документ
                    CreateEmployeesWordDocument(saveFileDialog.FileName, employeesData);
                    
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
        /// Получает данные сотрудников из базы данных
        /// </summary>
        private async Task<DataTable> GetEmployeesDataAsync()
        {
            var dt = new DataTable();
            await using var conn = new SqlConnection(ConnectionString);
            await conn.OpenAsync();
            string sql = @"SELECT EmployeeID, FirstName, LastName, Position, Email, Phone 
                          FROM dbo.Employees ORDER BY EmployeeID";
            await using var cmd = new SqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
            return dt;
        }

        /// <summary>
        /// Создаёт Word документ с таблицей сотрудников
        /// </summary>
        private void CreateEmployeesWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());

                // Заголовок
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Сотрудники"));

                // Таблица
                WordProcessing.Table table = CreateTable();

                // Заголовок таблицы
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Имя", true));
                headerRow.Append(CreateTableCell("Фамилия", true));
                headerRow.Append(CreateTableCell("Должность", true));
                headerRow.Append(CreateTableCell("Email", true));
                headerRow.Append(CreateTableCell("Телефон", true));
                table.Append(headerRow);

                // Данные
                foreach (DataRow row in data.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["EmployeeID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["FirstName"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["LastName"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Position"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Email"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Phone"]?.ToString() ?? "", false));
                    table.Append(dataRow);
                }

                body.Append(table);
                mainPart.Document.Save();
            }
        }

        // ===============================================
        // Экспорт в Word для Equipment
        // ===============================================

        /// <summary>
        /// Обработчик кнопки "Сохранить Word" для оборудования
        /// </summary>
        private async void SaveEquipmentWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "Отчет_Оборудование"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Получаем данные оборудования из БД
                    DataTable equipmentData = await GetEquipmentDataAsync();
                    
                    // Создаём Word документ
                    CreateEquipmentWordDocument(saveFileDialog.FileName, equipmentData);
                    
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
        /// Получает данные оборудования из базы данных
        /// </summary>
        private async Task<DataTable> GetEquipmentDataAsync()
        {
            var dt = new DataTable();
            await using var conn = new SqlConnection(ConnectionString);
            await conn.OpenAsync();
            string sql = @"SELECT e.EquipmentID, e.EquipmentName, c.CategoryName, e.Manufacturer, 
                          e.Model, e.SerialNumber, e.PurchaseDate, e.WarrantyUntil 
                          FROM dbo.Equipment e
                          LEFT JOIN dbo.Categories c ON e.CategoryID = c.CategoryID
                          ORDER BY e.EquipmentID";
            await using var cmd = new SqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
            return dt;
        }

        /// <summary>
        /// Создаёт Word документ с таблицей оборудования
        /// </summary>
        private void CreateEquipmentWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());

                // Заголовок
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Оборудование"));

                // Таблица
                WordProcessing.Table table = CreateTable();

                // Заголовок таблицы
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Название", true));
                headerRow.Append(CreateTableCell("Категория", true));
                headerRow.Append(CreateTableCell("Производитель", true));
                headerRow.Append(CreateTableCell("Модель", true));
                headerRow.Append(CreateTableCell("Серийный №", true));
                headerRow.Append(CreateTableCell("Дата покупки", true));
                headerRow.Append(CreateTableCell("Гарантия до", true));
                table.Append(headerRow);

                // Данные
                foreach (DataRow row in data.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["EquipmentID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["EquipmentName"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["CategoryName"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Manufacturer"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Model"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["SerialNumber"]?.ToString() ?? "", false));
                    // Форматируем даты
                    string purchaseDate = row["PurchaseDate"] != DBNull.Value 
                        ? ((DateTime)row["PurchaseDate"]).ToString("dd.MM.yyyy") : "";
                    string warrantyDate = row["WarrantyUntil"] != DBNull.Value 
                        ? ((DateTime)row["WarrantyUntil"]).ToString("dd.MM.yyyy") : "";
                    dataRow.Append(CreateTableCell(purchaseDate, false));
                    dataRow.Append(CreateTableCell(warrantyDate, false));
                    table.Append(dataRow);
                }

                body.Append(table);
                mainPart.Document.Save();
            }
        }

        // ===============================================
        // Экспорт в Word для Движений оборудования
        // ===============================================

        /// <summary>
        /// Обработчик кнопки "Сохранить Word" для движений оборудования
        /// </summary>
        private async void SaveMovementsWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "Отчет_Движения"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Получаем данные движений из БД
                    DataTable movementsData = await GetMovementsDataAsync();
                    
                    // Создаём Word документ
                    CreateMovementsWordDocument(saveFileDialog.FileName, movementsData);
                    
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
        /// Получает данные движений оборудования из базы данных
        /// </summary>
        private async Task<DataTable> GetMovementsDataAsync()
        {
            var dt = new DataTable();
            await using var conn = new SqlConnection(ConnectionString);
            await conn.OpenAsync();
            string sql = @"SELECT m.MovementID, e.EquipmentName, m.MovementDate, m.Quantity, 
                          m.MovementType, s.SupplierName, 
                          emp.FirstName + ' ' + emp.LastName AS EmployeeName, m.Notes
                          FROM dbo.EquipmentMovement m
                          LEFT JOIN dbo.Equipment e ON m.EquipmentID = e.EquipmentID
                          LEFT JOIN dbo.Suppliers s ON m.SupplierID = s.SupplierID
                          LEFT JOIN dbo.Employees emp ON m.EmployeeID = emp.EmployeeID
                          ORDER BY m.MovementID DESC";
            await using var cmd = new SqlCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
            return dt;
        }

        /// <summary>
        /// Создаёт Word документ с таблицей движений оборудования
        /// </summary>
        private void CreateMovementsWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());

                // Заголовок
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Движения оборудования"));

                // Таблица
                WordProcessing.Table table = CreateTable();

                // Заголовок таблицы
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Оборудование", true));
                headerRow.Append(CreateTableCell("Дата", true));
                headerRow.Append(CreateTableCell("Кол-во", true));
                headerRow.Append(CreateTableCell("Тип", true));
                headerRow.Append(CreateTableCell("Поставщик", true));
                headerRow.Append(CreateTableCell("Сотрудник", true));
                headerRow.Append(CreateTableCell("Примечания", true));
                table.Append(headerRow);

                // Данные
                foreach (DataRow row in data.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["MovementID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["EquipmentName"]?.ToString() ?? "", false));
                    // Форматируем дату
                    string movementDate = row["MovementDate"] != DBNull.Value 
                        ? ((DateTime)row["MovementDate"]).ToString("dd.MM.yyyy") : "";
                    dataRow.Append(CreateTableCell(movementDate, false));
                    dataRow.Append(CreateTableCell(row["Quantity"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["MovementType"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["SupplierName"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["EmployeeName"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Notes"]?.ToString() ?? "", false));
                    table.Append(dataRow);
                }

                body.Append(table);
                mainPart.Document.Save();
            }
        }

        // ===============================================
        // Вспомогательные методы для создания Word документов
        // ===============================================

        /// <summary>
        /// Создаёт параграф с заголовком документа
        /// </summary>
        private WordProcessing.Paragraph CreateTitleParagraph(string title)
        {
            WordProcessing.Paragraph paragraph = new WordProcessing.Paragraph();
            
            WordProcessing.ParagraphProperties props = new WordProcessing.ParagraphProperties();
            props.Append(new WordProcessing.Justification() { Val = WordProcessing.JustificationValues.Center });
            props.Append(new WordProcessing.SpacingBetweenLines() { After = "400" });
            paragraph.Append(props);

            WordProcessing.Run run = new WordProcessing.Run();
            WordProcessing.RunProperties runProps = new WordProcessing.RunProperties();
            runProps.Append(new WordProcessing.Bold());
            runProps.Append(new WordProcessing.FontSize() { Val = "56" }); // 28pt
            run.Append(runProps);
            run.Append(new WordProcessing.Text(title));
            
            paragraph.Append(run);
            return paragraph;
        }

        /// <summary>
        /// Создаёт параграф с подзаголовком
        /// </summary>
        private WordProcessing.Paragraph CreateSubtitleParagraph(string subtitle)
        {
            WordProcessing.Paragraph paragraph = new WordProcessing.Paragraph();
            
            WordProcessing.ParagraphProperties props = new WordProcessing.ParagraphProperties();
            props.Append(new WordProcessing.SpacingBetweenLines() { After = "200" });
            paragraph.Append(props);

            WordProcessing.Run run = new WordProcessing.Run();
            WordProcessing.RunProperties runProps = new WordProcessing.RunProperties();
            runProps.Append(new WordProcessing.Bold());
            runProps.Append(new WordProcessing.FontSize() { Val = "32" }); // 16pt
            run.Append(runProps);
            run.Append(new WordProcessing.Text(subtitle));
            
            paragraph.Append(run);
            return paragraph;
        }

        /// <summary>
        /// Создаёт таблицу Word с базовыми настройками (границы, ширина)
        /// </summary>
        private WordProcessing.Table CreateTable()
        {
            WordProcessing.Table table = new WordProcessing.Table();

            WordProcessing.TableProperties tableProps = new WordProcessing.TableProperties();
            WordProcessing.TableBorders tableBorders = new WordProcessing.TableBorders(
                new WordProcessing.TopBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                new WordProcessing.BottomBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                new WordProcessing.LeftBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                new WordProcessing.RightBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                new WordProcessing.InsideHorizontalBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 },
                new WordProcessing.InsideVerticalBorder() { Val = WordProcessing.BorderValues.Single, Size = 4 }
            );
            tableProps.Append(tableBorders);
            tableProps.Append(new WordProcessing.TableWidth() { Width = "5000", Type = WordProcessing.TableWidthUnitValues.Pct });
            table.Append(tableProps);

            return table;
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
