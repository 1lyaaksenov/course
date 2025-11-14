using System.Windows;

namespace course
{
    public partial class EquipmentDialog : Window
    {
        public Equipment Equipment { get; private set; }

        public EquipmentDialog() : this(new Equipment()) { }

        public EquipmentDialog(Equipment equipment)
        {
            InitializeComponent();
            Equipment = equipment;
            InitializeFields();
        }

        private void InitializeFields()
        {
            txtName.Text = Equipment.Name;
            cmbType.Text = Equipment.Type;
            txtSerialNumber.Text = Equipment.SerialNumber;
            dpPurchaseDate.SelectedDate = Equipment.PurchaseDate;
            cmbStatus.Text = Equipment.Status;
            txtLocation.Text = Equipment.Location;
            txtCost.Text = Equipment.Cost.ToString("F2");
            txtDescription.Text = Equipment.Description;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateInput())
            {
                Equipment.Name = txtName.Text.Trim();
                Equipment.Type = cmbType.Text;
                Equipment.SerialNumber = txtSerialNumber.Text.Trim();
                Equipment.PurchaseDate = dpPurchaseDate.SelectedDate ?? DateTime.Now;
                Equipment.Status = cmbStatus.Text;
                Equipment.Location = txtLocation.Text.Trim();
                Equipment.Cost = decimal.Parse(txtCost.Text);
                Equipment.Description = txtDescription.Text.Trim();

                DialogResult = true;
                Close();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Введите название оборудования", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(cmbType.Text))
            {
                MessageBox.Show("Выберите тип оборудования", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSerialNumber.Text))
            {
                MessageBox.Show("Введите серийный номер", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!decimal.TryParse(txtCost.Text, out decimal cost) || cost < 0)
            {
                MessageBox.Show("Введите корректную стоимость", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }
    }
}