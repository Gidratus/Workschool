using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;
namespace WpfApp9
{
    public partial class MyMovementsWindow : Window
    {
        private const string ConnectionString = App.ConnectionString;
        private readonly int _employeeId;
        public MyMovementsWindow(int employeeId)
        {
            InitializeComponent();
            _employeeId = employeeId;
            Loaded += MyMovementsWindow_Loaded;
        }
        private async void MyMovementsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadMovementsAsync();
        }
        private async Task LoadMovementsAsync()
        {
            try
            {
                var dt = new DataTable();
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                string sql = @"SELECT m.MovementID, m.Isdone, e.EquipmentName, m.MovementDate, m.Quantity, m.MovementType, 
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
        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dataView = MovementsDataGrid.ItemsSource as DataView;
                if (dataView == null) return;
                await using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();
                int updatedCount = 0;
                foreach (DataRowView row in dataView)
                {
                    int movementId = Convert.ToInt32(row["MovementID"]);
                    bool isDone = row["Isdone"] != DBNull.Value && Convert.ToBoolean(row["Isdone"]);
                    string sql = "UPDATE dbo.EquipmentMovement SET Isdone = @Isdone WHERE MovementID = @MovementID";
                    await using var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@Isdone", isDone);
                    cmd.Parameters.AddWithValue("@MovementID", movementId);
                    updatedCount += await cmd.ExecuteNonQueryAsync();
                }
                MessageBox.Show($"Изменения сохранены! Обновлено записей: {updatedCount}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
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
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
