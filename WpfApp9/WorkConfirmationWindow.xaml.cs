using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;

namespace WpfApp9
{
    /// <summary>
    /// Окно подтверждения выполненных работ для Менеджеров и Администраторов.
    /// Отображает записи, которые стажеры отметили как выполненные (Isdone = true),
    /// и позволяет менеджерам/админам подтвердить их.
    /// </summary>
    public partial class WorkConfirmationWindow : Window
    {
        // Строка подключения к базе данных
        private const string ConnectionString = App.ConnectionString;

        // Данные текущего пользователя (для проверки прав доступа)
        private readonly UserData _currentUser;

        /// <summary>
        /// Конструктор окна подтверждения работ
        /// </summary>
        /// <param name="userData">Данные текущего пользователя (ID и должность)</param>
        public WorkConfirmationWindow(UserData userData)
        {
            InitializeComponent();
            _currentUser = userData;
            Loaded += WorkConfirmationWindow_Loaded;
        }

        /// <summary>
        /// Обработчик загрузки окна - проверяет права и загружает данные
        /// </summary>
        private async void WorkConfirmationWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Проверка прав доступа - только Менеджер или Админ могут использовать это окно
                if (!IsUserAuthorized())
                {
                    MessageBox.Show("У вас нет прав для подтверждения работ. Доступ только для Менеджеров и Администраторов.",
                        "Доступ запрещен", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Close();
                    return;
                }

                // Загружаем записи, отмеченные стажерами как выполненные
                await LoadPendingConfirmationsAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке окна: {ex.Message}\n\nПодробности: {ex.StackTrace}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Проверка, имеет ли пользователь права на подтверждение работ
        /// Доступ только для Менеджеров и Администраторов
        /// </summary>
        /// <returns>true если пользователь имеет права, иначе false</returns>
        private bool IsUserAuthorized()
        {
            // Проверяем должность пользователя
            if (string.IsNullOrEmpty(_currentUser.Position))
                return false;

            // Разрешаем доступ для Менеджера, Админа, Администратора
            // Учитываем разные варианты написания (Менеджер/Менджер)
            string position = _currentUser.Position.ToLower();
            return position.Contains("менеджер") || 
                   position.Contains("менджер") ||    // вариант без "е"
                   position.Contains("админ") || 
                   position.Contains("администратор") ||
                   position.Contains("manager") ||
                   position.Contains("admin");
        }

        /// <summary>
        /// Загрузка записей, которые отмечены стажерами как выполненные (Isdone = 1)
        /// и ожидают подтверждения от менеджера
        /// </summary>
        private async Task LoadPendingConfirmationsAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                // SQL-запрос для получения записей, отмеченных как выполненные
                // Включаем информацию о сотруднике, который выполнил работу
                // isAdmindone - поле для подтверждения менеджером/админом
                string sql = @"SELECT m.MovementID, 
                                      m.Isdone, 
                                      ISNULL(m.isAdmindone, CAST(0 as bit)) as isAdmindone,
                                      e.EquipmentName, 
                                      (emp.FirstName + ' ' + emp.LastName) as EmployeeName,
                                      m.MovementDate, 
                                      m.Quantity, 
                                      m.MovementType, 
                                      s.SupplierName, 
                                      m.Notes
                              FROM dbo.EquipmentMovement m
                              LEFT JOIN dbo.Equipment e ON m.EquipmentID = e.EquipmentID
                              LEFT JOIN dbo.Suppliers s ON m.SupplierID = s.SupplierID
                              LEFT JOIN dbo.Employees emp ON m.EmployeeID = emp.EmployeeID
                              WHERE m.Isdone = 1
                              ORDER BY m.MovementDate DESC";

                await using var cmd = new SqlCommand(sql, conn);
                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);

                // Устанавливаем источник данных для DataGrid
                ConfirmationDataGrid.ItemsSource = dt.DefaultView;

                // Показываем информацию о количестве записей
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Нет работ, ожидающих подтверждения.",
                        "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при загрузке записей: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке записей: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Сохранение подтверждений в базу данных
        /// Обновляет поле IsConfirmed для выбранных записей
        /// </summary>
        private async void SaveConfirmationsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем DataView из DataGrid
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

                // Перебираем все строки и обновляем статус isAdmindone
                foreach (DataRowView row in dataView)
                {
                    int movementId = Convert.ToInt32(row["MovementID"]);
                    bool isAdminDone = row["isAdmindone"] != DBNull.Value && Convert.ToBoolean(row["isAdmindone"]);

                    // Обновляем только если подтверждено
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

                MessageBox.Show($"Подтверждения сохранены!\nПодтверждено работ: {confirmedCount}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                // Обновляем список после сохранения
                await LoadPendingConfirmationsAsync();
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"SQL ошибка при сохранении: {ex.Message}",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Обновление списка записей
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
