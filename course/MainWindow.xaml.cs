using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using System.Data.SqlClient;

namespace course
{
    public partial class MainWindow : Window
    {
        static MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private string _currentRole;
        private DatabaseService _dbService;
        private Equipment _selectedEquipment;
        private User _selectedUser;

        public MainWindow(string role, DatabaseService dbService)
        {

            InitializeComponent();
            _currentRole = role;
            _dbService = dbService;

            InitializeUI();
            LoadData();
        }

        // Остальной код без изменений...
        private void InitializeUI()
        {
            txtCurrentRole.Text = _currentRole;

            if (_currentRole == "User")
            {
                tabManagement.Visibility = Visibility.Collapsed;
                tabUsers.Visibility = Visibility.Collapsed;
            }
            else if (_currentRole == "Manager")
            {
                tabUsers.Visibility = Visibility.Collapsed;
            }

            UpdateStatus($"Добро пожаловать! Ваша роль: {_currentRole}");
        }

        private void LoadData()
        {
            try
            {
                var equipment = _dbService.GetAllEquipment();
                dgEquipment.ItemsSource = equipment;
                dgEquipmentManagement.ItemsSource = equipment;

                if (_currentRole == "Admin")
                {
                    var users = _dbService.GetAllUsers();
                    dgUsers.ItemsSource = users;
                }

                UpdateStatus("Данные успешно загружены");
            }
            catch (System.Exception ex)
            {
                UpdateStatus($"Ошибка загрузки данных: {ex.Message}");
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EquipmentSelection_Changed(object sender, SelectionChangedEventArgs e)
        {
            _selectedEquipment = dgEquipmentManagement.SelectedItem as Equipment;
        }

        private void UserSelection_Changed(object sender, SelectionChangedEventArgs e)
        {
            _selectedUser = dgUsers.SelectedItem as User;
        }

        private void AddEquipment_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new EquipmentDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    if (_dbService.AddEquipment(dialog.Equipment))
                    {
                        LoadData();
                        UpdateStatus("Оборудование успешно добавлено");
                        MessageBox.Show("Оборудование успешно добавлено", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка добавления оборудования: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void EditEquipment_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedEquipment == null)
            {
                MessageBox.Show("Выберите оборудование для редактирования", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new EquipmentDialog(_selectedEquipment);
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    if (_dbService.UpdateEquipment(dialog.Equipment))
                    {
                        LoadData();
                        UpdateStatus("Оборудование успешно обновлено");
                        MessageBox.Show("Оборудование успешно обновлено", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка обновления оборудования: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteEquipment_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedEquipment == null)
            {
                MessageBox.Show("Выберите оборудование для удаления", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show($"Вы уверены, что хотите удалить оборудование '{_selectedEquipment.Name}'?",
                "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (_dbService.DeleteEquipment(_selectedEquipment.Id))
                    {
                        LoadData();
                        UpdateStatus("Оборудование успешно удалено");
                        MessageBox.Show("Оборудование успешно удалено", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления оборудования: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void AddUser_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new UserDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    if (_dbService.AddUser(dialog.User))
                    {
                        LoadData();
                        UpdateStatus("Пользователь успешно добавлен");
                        MessageBox.Show("Пользователь успешно добавлен", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка добавления пользователя: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedUser == null)
            {
                MessageBox.Show("Выберите пользователя для удаления", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_selectedUser.Login == "admin")
            {
                MessageBox.Show("Нельзя удалить администратора системы", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var result = MessageBox.Show($"Вы уверены, что хотите удалить пользователя '{_selectedUser.Login}'?",
                "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (_dbService.DeleteUser(_selectedUser.Id))
                    {
                        LoadData();
                        UpdateStatus("Пользователь успешно удален");
                        MessageBox.Show("Пользователь успешно удален", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления пользователя: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void RefreshEquipment_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        private void RefreshUsers_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение выхода",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                var loginWindow = new LoginWindow();
                loginWindow.Show();
                this.Close();
            }
        }

        private void UpdateStatus(string message)
        {
            txtStatus.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        }

        // Методы экспорта в Excel (без установки лицензии внутри методов)
        private void ExportAllToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Отчет_по_оргтехнике_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    ExportAllDataToExcel(saveFileDialog.FileName);
                    UpdateStatus("Все данные успешно экспортированы в Excel");
                    MessageBox.Show("Все данные успешно экспортированы в Excel", "Экспорт завершен",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportEquipmentToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Оргтехника_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    ExportEquipmentToExcel(saveFileDialog.FileName);
                    UpdateStatus("Данные об оргтехнике экспортированы в Excel");
                    MessageBox.Show("Данные об оргтехнике успешно экспортированы в Excel", "Экспорт завершен",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportUsersToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Пользователи_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    ExportUsersToExcel(saveFileDialog.FileName);
                    UpdateStatus("Данные о пользователях экспортированы в Excel");
                    MessageBox.Show("Данные о пользователях успешно экспортированы в Excel", "Экспорт завершен",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportAllDataToExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                // Лист с оргтехникой
                var equipmentWorksheet = package.Workbook.Worksheets.Add("Оргтехника");
                equipmentWorksheet.Cells[1, 1].Value = "Отчет по оргтехнике";
                equipmentWorksheet.Cells[1, 1, 1, 8].Merge = true;
                equipmentWorksheet.Cells[1, 1].Style.Font.Bold = true;
                equipmentWorksheet.Cells[1, 1].Style.Font.Size = 14;
                equipmentWorksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                // Заголовки для оргтехники
                string[] equipmentHeaders = { "Название", "Тип", "Серийный номер", "Дата покупки", "Статус", "Местоположение", "Стоимость", "Описание" };
                for (int i = 0; i < equipmentHeaders.Length; i++)
                {
                    equipmentWorksheet.Cells[3, i + 1].Value = equipmentHeaders[i];
                    equipmentWorksheet.Cells[3, i + 1].Style.Font.Bold = true;
                    equipmentWorksheet.Cells[3, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    equipmentWorksheet.Cells[3, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                var equipment = _dbService.GetAllEquipment();
                for (int i = 0; i < equipment.Count; i++)
                {
                    var item = equipment[i];
                    equipmentWorksheet.Cells[i + 4, 1].Value = item.Name;
                    equipmentWorksheet.Cells[i + 4, 2].Value = item.Type;
                    equipmentWorksheet.Cells[i + 4, 3].Value = item.SerialNumber;
                    equipmentWorksheet.Cells[i + 4, 4].Value = item.PurchaseDate.ToString("dd.MM.yyyy");
                    equipmentWorksheet.Cells[i + 4, 5].Value = item.Status;
                    equipmentWorksheet.Cells[i + 4, 6].Value = item.Location;
                    equipmentWorksheet.Cells[i + 4, 7].Value = item.Cost;
                    equipmentWorksheet.Cells[i + 4, 8].Value = item.Description;
                }

                equipmentWorksheet.Cells[equipmentWorksheet.Dimension.Address].AutoFitColumns();

                // Лист с пользователями (только для админов)
                if (_currentRole == "Admin")
                {
                    var usersWorksheet = package.Workbook.Worksheets.Add("Пользователи");
                    usersWorksheet.Cells[1, 1].Value = "Пользователи системы";
                    usersWorksheet.Cells[1, 1, 1, 3].Merge = true;
                    usersWorksheet.Cells[1, 1].Style.Font.Bold = true;
                    usersWorksheet.Cells[1, 1].Style.Font.Size = 14;
                    usersWorksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    string[] userHeaders = { "Логин", "Пароль", "Роль" };
                    for (int i = 0; i < userHeaders.Length; i++)
                    {
                        usersWorksheet.Cells[3, i + 1].Value = userHeaders[i];
                        usersWorksheet.Cells[3, i + 1].Style.Font.Bold = true;
                        usersWorksheet.Cells[3, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        usersWorksheet.Cells[3, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    var users = _dbService.GetAllUsers();
                    for (int i = 0; i < users.Count; i++)
                    {
                        var user = users[i];
                        usersWorksheet.Cells[i + 4, 1].Value = user.Login;
                        usersWorksheet.Cells[i + 4, 2].Value = user.Password;
                        usersWorksheet.Cells[i + 4, 3].Value = user.Role;
                    }

                    usersWorksheet.Cells[usersWorksheet.Dimension.Address].AutoFitColumns();
                }

                // Лист с итогами
                var summaryWorksheet = package.Workbook.Worksheets.Add("Итоги");
                summaryWorksheet.Cells[1, 1].Value = "Сводный отчет";
                summaryWorksheet.Cells[1, 1, 1, 3].Merge = true;
                summaryWorksheet.Cells[1, 1].Style.Font.Bold = true;
                summaryWorksheet.Cells[1, 1].Style.Font.Size = 14;
                summaryWorksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                summaryWorksheet.Cells[3, 1].Value = "Общее количество единиц техники:";
                summaryWorksheet.Cells[3, 2].Value = equipment.Count;
                summaryWorksheet.Cells[4, 1].Value = "Общая стоимость техники:";
                summaryWorksheet.Cells[4, 2].Value = equipment.Sum(e => e.Cost);
                summaryWorksheet.Cells[4, 2].Style.Numberformat.Format = "#,##0.00\"р.\"";

                // Статистика по статусам
                summaryWorksheet.Cells[6, 1].Value = "Статистика по статусам:";
                summaryWorksheet.Cells[6, 1].Style.Font.Bold = true;

                var statusGroups = equipment.GroupBy(e => e.Status)
                                          .Select(g => new { Status = g.Key, Count = g.Count() })
                                          .ToList();

                for (int i = 0; i < statusGroups.Count; i++)
                {
                    summaryWorksheet.Cells[i + 7, 1].Value = statusGroups[i].Status;
                    summaryWorksheet.Cells[i + 7, 2].Value = statusGroups[i].Count;
                }

                // Статистика по типам
                summaryWorksheet.Cells[6, 4].Value = "Статистика по типам:";
                summaryWorksheet.Cells[6, 4].Style.Font.Bold = true;

                var typeGroups = equipment.GroupBy(e => e.Type)
                                         .Select(g => new { Type = g.Key, Count = g.Count() })
                                         .ToList();

                for (int i = 0; i < typeGroups.Count; i++)
                {
                    summaryWorksheet.Cells[i + 7, 4].Value = typeGroups[i].Type;
                    summaryWorksheet.Cells[i + 7, 5].Value = typeGroups[i].Count;
                }

                summaryWorksheet.Cells[summaryWorksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void ExportEquipmentToExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Оргтехника");

                // Заголовок
                worksheet.Cells[1, 1].Value = "Отчет по оргтехнике";
                worksheet.Cells[1, 1, 1, 8].Merge = true;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Style.Font.Size = 14;
                worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                // Заголовки столбцов
                string[] headers = { "Название", "Тип", "Серийный номер", "Дата покупки", "Статус", "Местоположение", "Стоимость", "Описание" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[3, i + 1].Value = headers[i];
                    worksheet.Cells[3, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[3, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[3, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                var equipment = _dbService.GetAllEquipment();
                for (int i = 0; i < equipment.Count; i++)
                {
                    var item = equipment[i];
                    worksheet.Cells[i + 4, 1].Value = item.Name;
                    worksheet.Cells[i + 4, 2].Value = item.Type;
                    worksheet.Cells[i + 4, 3].Value = item.SerialNumber;
                    worksheet.Cells[i + 4, 4].Value = item.PurchaseDate.ToString("dd.MM.yyyy");
                    worksheet.Cells[i + 4, 5].Value = item.Status;
                    worksheet.Cells[i + 4, 6].Value = item.Location;
                    worksheet.Cells[i + 4, 7].Value = item.Cost;
                    worksheet.Cells[i + 4, 8].Value = item.Description;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void ExportUsersToExcel(string filePath)
        {
            if (_currentRole != "Admin")
            {
                MessageBox.Show("Экспорт данных о пользователях доступен только администраторам", "Доступ запрещен",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Пользователи");

                // Заголовок
                worksheet.Cells[1, 1].Value = "Пользователи системы";
                worksheet.Cells[1, 1, 1, 3].Merge = true;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Style.Font.Size = 14;
                worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                // Заголовки столбцов
                string[] headers = { "Логин", "Пароль", "Роль" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[3, i + 1].Value = headers[i];
                    worksheet.Cells[3, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[3, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[3, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                var users = _dbService.GetAllUsers();
                for (int i = 0; i < users.Count; i++)
                {
                    var user = users[i];
                    worksheet.Cells[i + 4, 1].Value = user.Login;
                    worksheet.Cells[i + 4, 2].Value = user.Password;
                    worksheet.Cells[i + 4, 3].Value = user.Role;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}