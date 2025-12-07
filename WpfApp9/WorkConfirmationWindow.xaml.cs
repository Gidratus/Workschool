using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;

namespace WpfApp9
{
    /// <summary>
    /// Окно подтверждения выполненных работ для Менеджеров и Администраторов
    /// </summary>
    public partial class WorkConfirmationWindow : Window
    {
        // Строка подключения к базе данных
        private const string ConnectionString = App.ConnectionString;

        // ID текущего пользователя
        private readonly int _employeeId;
        
        // Должность текущего пользователя
        private string _position;

        /// <summary>
        /// Конструктор окна подтверждения работ
        /// </summary>
        /// <param name="employeeId">ID пользователя</param>
        public WorkConfirmationWindow(int employeeId)
        {
            InitializeComponent();
            _employeeId = employeeId;
            Loaded += WorkConfirmationWindow_Loaded;
        }

        /// <summary>
        /// Обработчик загрузки окна - проверяет права и загружает данные
        /// </summary>
        private async void WorkConfirmationWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем должность пользователя из БД
                await LoadUserPositionAsync();
                
                // Проверка прав доступа - только Менеджер или Админ
                if (!IsUserAuthorized())
                {
                    MessageBox.Show("У вас нет прав для подтверждения работ.\nДоступ только для Менеджеров и Администраторов.",
                        "Доступ запрещен", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Close();
                    return;
                }

                // Загружаем записи, отмеченные стажерами как выполненные
                await LoadPendingConfirmationsAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Загрузка должности пользователя из базы данных
        /// </summary>
        private async Task LoadUserPositionAsync()
        {
            try
            {
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                string sql = "SELECT Position FROM dbo.Employees WHERE EmployeeID = @EmployeeID";
                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", _employeeId);

                var result = await cmd.ExecuteScalarAsync();
                _position = result?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных пользователя: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                _position = "";
            }
        }

        /// <summary>
        /// Проверка прав доступа (Менеджер или Админ)
        /// </summary>
        private bool IsUserAuthorized()
        {
            if (string.IsNullOrEmpty(_position))
                return false;

            string position = _position.ToLower();
            return position.Contains("менеджер") || 
                   position.Contains("менджер") ||
                   position.Contains("админ") || 
                   position.Contains("администратор");
        }

        /// <summary>
        /// Загрузка записей, отмеченных стажерами как выполненные
        /// </summary>
        private async Task LoadPendingConfirmationsAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                // SQL-запрос для получения записей с Isdone = 1
                string sql = @"SELECT m.MovementID, 
                                      m.Isdone, 
                                      m.isAdmindone,
                                      e.EquipmentName, 
                                      (emp.FirstName + ' ' + emp.LastName) as EmployeeName,
                                      m.MovementDate, 
                                      m.Quantity, 
                                      m.MovementType, 
                                      m.Notes
                              FROM dbo.EquipmentMovement m
                              LEFT JOIN dbo.Equipment e ON m.EquipmentID = e.EquipmentID
                              LEFT JOIN dbo.Employees emp ON m.EmployeeID = emp.EmployeeID
                              WHERE m.Isdone = 1
                              ORDER BY m.MovementDate DESC";

                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);

                ConfirmationDataGrid.ItemsSource = dt.DefaultView;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Нет работ, ожидающих подтверждения.",
                        "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Сохранение подтверждений
        /// </summary>
        private async void SaveConfirmationsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dataView = ConfirmationDataGrid.ItemsSource as DataView;
                if (dataView == null || dataView.Count == 0)
                {
                    MessageBox.Show("Нет данных для сохранения.",
                        "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                int confirmedCount = 0;

                foreach (DataRowView row in dataView)
                {
                    int movementId = Convert.ToInt32(row["MovementID"]);
                    bool isAdminDone = row["isAdmindone"] != DBNull.Value && Convert.ToBoolean(row["isAdmindone"]);

                    if (isAdminDone)
                    {
                        string sql = "UPDATE dbo.EquipmentMovement SET isAdmindone = @isAdmindone WHERE MovementID = @MovementID";
                        await using var cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@isAdmindone", isAdminDone);
                        cmd.Parameters.AddWithValue("@MovementID", movementId);

                        await cmd.ExecuteNonQueryAsync();
                        confirmedCount++;
                    }
                }

                MessageBox.Show($"Подтверждено работ: {confirmedCount}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                await LoadPendingConfirmationsAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обновление списка
        /// </summary>
        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await LoadPendingConfirmationsAsync();
        }

        /// <summary>
        /// Закрытие окна
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
