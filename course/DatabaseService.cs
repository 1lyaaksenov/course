using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace course
{
    public class DatabaseService
    {
        private readonly string _connectionString = @"Data Source=stud-mssql.sttec.yar.ru,38325;Initial Catalog=user12_db;User ID=user12_db;Password=user12";

        public User Authenticate(string username, string password)
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = @"
                SELECT u.Id, u.Username, u.Password, u.RoleId, u.FullName, r.Name as RoleName 
                FROM Users u 
                INNER JOIN Roles r ON u.RoleId = r.Id 
                WHERE u.Username = @Username AND u.Password = @Password";

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Username", username);
            command.Parameters.AddWithValue("@Password", password);

            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                return new User
                {
                    Id = reader.GetInt32("Id"),
                    Login = reader.GetString("Username"),
                    Password = reader.GetString("Password"),
                    Role = reader.GetString("RoleName")
                };
            }
            return null;
        }

        public void InitializeDatabase()
        {
            // База данных создается через SQL скрипт
        }

        public List<Equipment> GetAllEquipment()
        {
            var equipment = new List<Equipment>();

            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = "SELECT * FROM Equipment ORDER BY Name";
            using var command = new SqlCommand(query, connection);
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                equipment.Add(new Equipment
                {
                    Id = reader.GetInt32("Id"),
                    Name = reader.GetString("Name"),
                    Type = reader.GetString("Type"),
                    SerialNumber = reader.GetString("SerialNumber"),
                    PurchaseDate = reader.GetDateTime("PurchaseDate"),
                    Status = reader.GetString("Status"),
                    Location = reader.GetString("Location"),
                    Cost = reader.GetDecimal("Cost"),
                    Description = reader.IsDBNull("Description") ? "" : reader.GetString("Description")
                });
            }

            return equipment;
        }

        public List<User> GetAllUsers()
        {
            var users = new List<User>();

            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = @"
                SELECT u.Id, u.Username, u.Password, u.RoleId, u.FullName, r.Name as RoleName
                FROM Users u 
                INNER JOIN Roles r ON u.RoleId = r.Id 
                ORDER BY u.Username";

            using var command = new SqlCommand(query, connection);
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                users.Add(new User
                {
                    Id = reader.GetInt32("Id"),
                    Login = reader.GetString("Username"),
                    Password = reader.GetString("Password"),
                    Role = reader.GetString("RoleName")
                });
            }

            return users;
        }

        public bool AddEquipment(Equipment equipment)
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();

                var query = @"INSERT INTO Equipment (Name, Type, SerialNumber, PurchaseDate, Status, Location, Cost, Description)
                             VALUES (@Name, @Type, @SerialNumber, @PurchaseDate, @Status, @Location, @Cost, @Description)";

                using var command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Name", equipment.Name);
                command.Parameters.AddWithValue("@Type", equipment.Type);
                command.Parameters.AddWithValue("@SerialNumber", equipment.SerialNumber);
                command.Parameters.AddWithValue("@PurchaseDate", equipment.PurchaseDate);
                command.Parameters.AddWithValue("@Status", equipment.Status);
                command.Parameters.AddWithValue("@Location", equipment.Location);
                command.Parameters.AddWithValue("@Cost", equipment.Cost);
                command.Parameters.AddWithValue("@Description", equipment.Description);

                return command.ExecuteNonQuery() > 0;
            }
            catch (SqlException ex)
            {
                if (ex.Number == 2627) // Ошибка уникальности
                    throw new Exception("Оборудование с таким серийным номером уже существует");
                throw;
            }
        }

        public bool UpdateEquipment(Equipment equipment)
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = @"UPDATE Equipment SET 
                         Name = @Name, Type = @Type, SerialNumber = @SerialNumber, 
                         PurchaseDate = @PurchaseDate, Status = @Status, Location = @Location,
                         Cost = @Cost, Description = @Description
                         WHERE Id = @Id";

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", equipment.Id);
            command.Parameters.AddWithValue("@Name", equipment.Name);
            command.Parameters.AddWithValue("@Type", equipment.Type);
            command.Parameters.AddWithValue("@SerialNumber", equipment.SerialNumber);
            command.Parameters.AddWithValue("@PurchaseDate", equipment.PurchaseDate);
            command.Parameters.AddWithValue("@Status", equipment.Status);
            command.Parameters.AddWithValue("@Location", equipment.Location);
            command.Parameters.AddWithValue("@Cost", equipment.Cost);
            command.Parameters.AddWithValue("@Description", equipment.Description);

            return command.ExecuteNonQuery() > 0;
        }

        public bool DeleteEquipment(int id)
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = "DELETE FROM Equipment WHERE Id = @Id";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", id);

            return command.ExecuteNonQuery() > 0;
        }

        public bool AddUser(User user)
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            // Получаем RoleId по названию роли
            var getRoleIdQuery = "SELECT Id FROM Roles WHERE Name = @RoleName";
            using var roleCommand = new SqlCommand(getRoleIdQuery, connection);
            roleCommand.Parameters.AddWithValue("@RoleName", user.Role);
            var roleId = roleCommand.ExecuteScalar();

            if (roleId == null)
            {
                throw new Exception($"Роль '{user.Role}' не найдена");
            }

            var query = "INSERT INTO Users (Username, Password, RoleId, FullName) VALUES (@Username, @Password, @RoleId, @FullName)";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Username", user.Login);
            command.Parameters.AddWithValue("@Password", user.Password);
            command.Parameters.AddWithValue("@RoleId", roleId);
            command.Parameters.AddWithValue("@FullName", user.Login);

            return command.ExecuteNonQuery() > 0;
        }

        public bool DeleteUser(int id)
        {
            using var connection = new SqlConnection(_connectionString);
            connection.Open();

            var query = "DELETE FROM Users WHERE Id = @Id";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", id);

            return command.ExecuteNonQuery() > 0;
        }
    }
}