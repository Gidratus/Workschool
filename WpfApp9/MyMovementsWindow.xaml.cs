using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;

namespace WpfApp9
{
    /// <summary>
    /// Окно просмотра записей движений оборудования для текущего пользователя
    /// </summary>
    public partial class MyMovementsWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;

        // ID текущего пользователя
        private readonly int _employeeId;

        /// <summary>
        /// Конструктор окна просмотра записей
        /// </summary>
        /// <param name="employeeId">ID пользователя для фильтрации записей</param>
        public MyMovementsWindow(int employeeId)
        {
            InitializeComponent();
            _employeeId = employeeId;
            Loaded += MyMovementsWindow_Loaded;
        }

        /// <summary>
        /// Загрузка данных при открытии окна
        /// </summary>
        private async void MyMovementsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadMovementsAsync();
        }

        /// <summary>
        /// Загрузка записей движений оборудования для текущего пользователя
        /// </summary>
        private async Task LoadMovementsAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                // SQL-запрос для получения записей движений с JOIN для имен оборудования и поставщика
                // Фильтруем по EmployeeID текущего пользователя
                string sql = @"SELECT e.EquipmentName, m.MovementDate, m.Quantity, m.MovementType, 
                              s.SupplierName, m.Notes
                              FROM dbo.EquipmentMovement m
                              LEFT JOIN dbo.Equipment e ON m.EquipmentID = e.EquipmentID
                              LEFT JOIN dbo.Suppliers s ON m.SupplierID = s.SupplierID
                              WHERE m.EmployeeID = @EmployeeID
                              ORDER BY m.MovementDate DESC";

                await using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", _employeeId);

                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);
                MovementsDataGrid.ItemsSource = dt.DefaultView;
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
        /// Закрытие окна
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

