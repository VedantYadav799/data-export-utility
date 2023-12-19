using System.Data;
using MySql.Data.MySqlClient;

namespace DEU;
public static class MyDAL
{
    public static DataTable GetData()
    {
        string connectionString = "server=localhost; database=inventorymanagement1; user=root; password=1234512345";
        string query = "SELECT * FROM employees";

        using (MySqlConnection connection = new(connectionString))
        {
            using (MySqlCommand cmd = new(query, connection))
            {
                DataTable dataTable = new();
                connection.Open();
                using MySqlDataAdapter adapter = new(cmd);
                adapter.Fill(dataTable);
                return dataTable;
            }
        }

    }
}

