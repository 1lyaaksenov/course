using System.Windows;
using System.Windows.Input;

namespace course
{
    public partial class LoginWindow : Window
    {
        private DatabaseService _dbService;

        public LoginWindow()
        {
            InitializeComponent();
            _dbService = new DatabaseService();

            // Инициализируем базу данных при запуске
            try
            {
                _dbService.InitializeDatabase();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации базы данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Устанавливаем фокус на поле логина
            Loaded += (s, e) => txtLogin.Focus();

            // Обработка нажатия Enter
            txtPassword.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter)
                    Login_Click(s, e);
            };
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Password;

            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Введите логин и пароль", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var user = _dbService.Authenticate(login, password);

                if (user != null)
                {
                    var mainWindow = new MainWindow(user.Role, _dbService);
                    mainWindow.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Неверные учетные данные", "Ошибка авторизации",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    txtPassword.Clear();
                    txtPassword.Focus();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к базе данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}