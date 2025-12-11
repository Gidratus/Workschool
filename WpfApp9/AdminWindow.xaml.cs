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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WordProcessing = DocumentFormat.OpenXml.Wordprocessing;
namespace WpfApp9
{
    public partial class AdminWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;
        private readonly int _currentEmployeeId;
        public AdminWindow(int employeeId)
        {
            InitializeComponent();
            _currentEmployeeId = employeeId;
            Loaded += AdminWindow_Loaded;
        }
        private async void AdminWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (_currentEmployeeId != 1)
            {
                EmployeesTabItem.Visibility = Visibility.Collapsed;
            }
            await LoadCategoriesAsync();
            await LoadCategoriesForEquipmentComboBoxAsync(); 
            await LoadEquipmentAsync(); 
            await LoadEquipmentForMovementComboBoxAsync();
            await LoadSuppliersForMovementComboBoxAsync();
            await LoadEmployeesForMovementComboBoxAsync();
            await LoadMovementAsync();
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
        private async void SaveCategoriesWordButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "Отчет_Категории"
                };
                if (saveFileDialog.ShowDialog() == true)
                {
                    DataTable categoriesData = await GetCategoriesDataAsync();
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
        private void CreateWordDocumentWithCategories(string filePath, DataTable categoriesData)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());
                WordProcessing.Paragraph titleParagraph = new WordProcessing.Paragraph();
                WordProcessing.ParagraphProperties titleParagraphProps = new WordProcessing.ParagraphProperties();
                titleParagraphProps.Append(new WordProcessing.Justification() { Val = WordProcessing.JustificationValues.Center });
                titleParagraphProps.Append(new WordProcessing.SpacingBetweenLines() { After = "400" });
                titleParagraph.Append(titleParagraphProps);
                WordProcessing.Run titleRun = new WordProcessing.Run();
                WordProcessing.RunProperties titleRunProps = new WordProcessing.RunProperties();
                titleRunProps.Append(new WordProcessing.Bold());
                titleRunProps.Append(new WordProcessing.FontSize() { Val = "56" }); 
                titleRun.Append(titleRunProps);
                titleRun.Append(new WordProcessing.Text("Отчет"));
                titleParagraph.Append(titleRun);
                body.Append(titleParagraph);
                WordProcessing.Paragraph subtitleParagraph = new WordProcessing.Paragraph();
                WordProcessing.ParagraphProperties subtitleProps = new WordProcessing.ParagraphProperties();
                subtitleProps.Append(new WordProcessing.SpacingBetweenLines() { After = "200" });
                subtitleParagraph.Append(subtitleProps);
                WordProcessing.Run subtitleRun = new WordProcessing.Run();
                WordProcessing.RunProperties subtitleRunProps = new WordProcessing.RunProperties();
                subtitleRunProps.Append(new WordProcessing.Bold());
                subtitleRunProps.Append(new WordProcessing.FontSize() { Val = "32" }); 
                subtitleRun.Append(subtitleRunProps);
                subtitleRun.Append(new WordProcessing.Text("Категории"));
                subtitleParagraph.Append(subtitleRun);
                body.Append(subtitleParagraph);
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
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Название категории", true));
                table.Append(headerRow);
                foreach (DataRow row in categoriesData.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["CategoryID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["CategoryName"].ToString() ?? "", false));
                    table.Append(dataRow);
                }
                body.Append(table);
                mainPart.Document.Save();
            }
        }
        private WordProcessing.TableCell CreateTableCell(string text, bool isHeader)
        {
            WordProcessing.TableCell cell = new WordProcessing.TableCell();
            WordProcessing.TableCellProperties cellProps = new WordProcessing.TableCellProperties();
            cellProps.Append(new WordProcessing.TableCellMargin(
                new WordProcessing.TopMargin() { Width = "50", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.BottomMargin() { Width = "50", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.LeftMargin() { Width = "100", Type = WordProcessing.TableWidthUnitValues.Dxa },
                new WordProcessing.RightMargin() { Width = "100", Type = WordProcessing.TableWidthUnitValues.Dxa }
            ));
            if (isHeader)
            {
                cellProps.Append(new WordProcessing.Shading() 
                { 
                    Val = WordProcessing.ShadingPatternValues.Clear, 
                    Fill = "DDDDDD" 
                });
            }
            cell.Append(cellProps);
            WordProcessing.Paragraph paragraph = new WordProcessing.Paragraph();
            WordProcessing.Run run = new WordProcessing.Run();
            WordProcessing.RunProperties runProps = new WordProcessing.RunProperties();
            runProps.Append(new WordProcessing.FontSize() { Val = "24" }); 
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
                    DataTable employeesData = await GetEmployeesDataAsync();
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
        private void CreateEmployeesWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Сотрудники"));
                WordProcessing.Table table = CreateTable();
                WordProcessing.TableRow headerRow = new WordProcessing.TableRow();
                headerRow.Append(CreateTableCell("ID", true));
                headerRow.Append(CreateTableCell("Имя", true));
                headerRow.Append(CreateTableCell("Фамилия", true));
                headerRow.Append(CreateTableCell("Должность", true));
                headerRow.Append(CreateTableCell("Email", true));
                headerRow.Append(CreateTableCell("Телефон", true));
                table.Append(headerRow);
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
                    DataTable equipmentData = await GetEquipmentDataAsync();
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
        private void CreateEquipmentWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Оборудование"));
                WordProcessing.Table table = CreateTable();
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
                foreach (DataRow row in data.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["EquipmentID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["EquipmentName"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["CategoryName"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Manufacturer"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["Model"]?.ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["SerialNumber"]?.ToString() ?? "", false));
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
                    DataTable movementsData = await GetMovementsDataAsync();
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
        private void CreateMovementsWordDocument(string filePath, DataTable data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new WordProcessing.Document();
                WordProcessing.Body body = mainPart.Document.AppendChild(new WordProcessing.Body());
                body.Append(CreateTitleParagraph("Отчет"));
                body.Append(CreateSubtitleParagraph("Движения оборудования"));
                WordProcessing.Table table = CreateTable();
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
                foreach (DataRow row in data.Rows)
                {
                    WordProcessing.TableRow dataRow = new WordProcessing.TableRow();
                    dataRow.Append(CreateTableCell(row["MovementID"].ToString() ?? "", false));
                    dataRow.Append(CreateTableCell(row["EquipmentName"]?.ToString() ?? "", false));
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
            runProps.Append(new WordProcessing.FontSize() { Val = "56" }); 
            run.Append(runProps);
            run.Append(new WordProcessing.Text(title));
            paragraph.Append(run);
            return paragraph;
        }
        private WordProcessing.Paragraph CreateSubtitleParagraph(string subtitle)
        {
            WordProcessing.Paragraph paragraph = new WordProcessing.Paragraph();
            WordProcessing.ParagraphProperties props = new WordProcessing.ParagraphProperties();
            props.Append(new WordProcessing.SpacingBetweenLines() { After = "200" });
            paragraph.Append(props);
            WordProcessing.Run run = new WordProcessing.Run();
            WordProcessing.RunProperties runProps = new WordProcessing.RunProperties();
            runProps.Append(new WordProcessing.Bold());
            runProps.Append(new WordProcessing.FontSize() { Val = "32" }); 
            run.Append(runProps);
            run.Append(new WordProcessing.Text(subtitle));
            paragraph.Append(run);
            return paragraph;
        }
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
        private async void AddEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
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
        private void ClearEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            ClearEmployeeFields();
            EmployeesDataGrid.SelectedItem = null;
        }
        private async void RefreshEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadEmployeesAsync();
            MessageBox.Show("Список сотрудников обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
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
        private async Task LoadEquipmentAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
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
        private void EquipmentDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EquipmentDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)EquipmentDataGrid.SelectedItem;
                EquipmentIdTextBox.Text = row["EquipmentID"].ToString();
                EquipmentNameTextBox.Text = row["EquipmentName"].ToString();
                EquipmentCategoryComboBox.SelectedValue = row["CategoryID"] != DBNull.Value ? row["CategoryID"] : null;
                EquipmentManufacturerTextBox.Text = row["Manufacturer"] != DBNull.Value ? row["Manufacturer"].ToString() : "";
                EquipmentModelTextBox.Text = row["Model"] != DBNull.Value ? row["Model"].ToString() : "";
                EquipmentSerialNumberTextBox.Text = row["SerialNumber"] != DBNull.Value ? row["SerialNumber"].ToString() : "";
                EquipmentPurchaseDatePicker.SelectedDate = row["PurchaseDate"] != DBNull.Value 
                    ? (DateTime?)row["PurchaseDate"] : null;
                EquipmentWarrantyUntilPicker.SelectedDate = row["WarrantyUntil"] != DBNull.Value 
                    ? (DateTime?)row["WarrantyUntil"] : null;
            }
        }
        private async void AddEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
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
                string sql = @"INSERT INTO dbo.Equipment 
                              (EquipmentName, CategoryID, Manufacturer, Model, SerialNumber, PurchaseDate, WarrantyUntil) 
                              VALUES (@EquipmentName, @CategoryID, @Manufacturer, @Model, @SerialNumber, @PurchaseDate, @WarrantyUntil)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentName", EquipmentNameTextBox.Text.Trim());
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
        private async void UpdateEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EquipmentIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите оборудование для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
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
        private async void DeleteEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EquipmentIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите оборудование для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
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
        private void ClearEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            ClearEquipmentFields();
            EquipmentDataGrid.SelectedItem = null;
        }
        private async void RefreshEquipmentButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadEquipmentAsync();
            MessageBox.Show("Список оборудования обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void ClearEquipmentFields()
        {
            EquipmentIdTextBox.Clear();
            EquipmentNameTextBox.Clear();
            EquipmentCategoryComboBox.SelectedIndex = -1; 
            EquipmentManufacturerTextBox.Clear();
            EquipmentModelTextBox.Clear();
            EquipmentSerialNumberTextBox.Clear();
            EquipmentPurchaseDatePicker.SelectedDate = null;
            EquipmentWarrantyUntilPicker.SelectedDate = null;
        }
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
        private async Task LoadEmployeesForMovementComboBoxAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
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
        private async Task LoadMovementAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
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
        private void MovementDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MovementDataGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)MovementDataGrid.SelectedItem;
                MovementIdTextBox.Text = row["MovementID"].ToString();
                MovementEquipmentComboBox.SelectedValue = row["EquipmentID"] != DBNull.Value ? row["EquipmentID"] : null;
                MovementDatePicker.SelectedDate = row["MovementDate"] != DBNull.Value 
                    ? (DateTime?)row["MovementDate"] : null;
                MovementQuantityTextBox.Text = row["Quantity"] != DBNull.Value ? row["Quantity"].ToString() : "";
                string movementType = row["MovementType"]?.ToString() ?? "";
                foreach (ComboBoxItem item in MovementTypeComboBox.Items)
                {
                    if (item.Content?.ToString() == movementType)
                    {
                        MovementTypeComboBox.SelectedItem = item;
                        break;
                    }
                }
                MovementSupplierComboBox.SelectedValue = row["SupplierID"] != DBNull.Value ? row["SupplierID"] : null;
                MovementEmployeeComboBox.SelectedValue = row["EmployeeID"] != DBNull.Value ? row["EmployeeID"] : null;
                MovementNotesTextBox.Text = row["Notes"] != DBNull.Value ? row["Notes"].ToString() : "";
            }
        }
        private async void AddMovementButton_Click(object sender, RoutedEventArgs e)
        {
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
                string sql = @"INSERT INTO dbo.EquipmentMovement 
                              (EquipmentID, MovementDate, Quantity, MovementType, SupplierID, Notes, EmployeeID) 
                              VALUES (@EquipmentID, @MovementDate, @Quantity, @MovementType, @SupplierID, @Notes, @EmployeeID)";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EquipmentID", MovementEquipmentComboBox.SelectedValue);
                cmd.Parameters.AddWithValue("@MovementDate", MovementDatePicker.SelectedDate.Value);
                cmd.Parameters.AddWithValue("@Quantity", quantity.HasValue ? (object)quantity.Value : DBNull.Value);
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
        private async void UpdateMovementButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(MovementIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите движение для обновления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
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
        private async void DeleteMovementButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(MovementIdTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, выберите движение для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
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
        private void ClearMovementButton_Click(object sender, RoutedEventArgs e)
        {
            ClearMovementFields();
            MovementDataGrid.SelectedItem = null;
        }
        private async void RefreshMovementButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadEquipmentForMovementComboBoxAsync();
            await LoadSuppliersForMovementComboBoxAsync();
            await LoadEmployeesForMovementComboBoxAsync();
            await LoadMovementAsync();
            MessageBox.Show("Список движений оборудования обновлен!",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
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
