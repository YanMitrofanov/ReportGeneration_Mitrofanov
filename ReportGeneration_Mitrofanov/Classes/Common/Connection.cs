using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace ReportGeneration_Mitrofanov.Classes.Common
{
    public class Connection
    {
        private static string connectionString = "server=localhost;port=3306;database=journal;uid=root;pwd=;";

        public static MySqlConnection OpenConnection()
        {
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            return connection;
        }

        public static void CloseConnection(MySqlConnection connection)
        {
            if (connection != null && connection.State == System.Data.ConnectionState.Open)
                connection.Close();
        }

        public static MySqlDataReader Query(string sql, MySqlConnection connection)
        {
            MySqlCommand command = new MySqlCommand(sql, connection);
            return command.ExecuteReader();
        }
    }
}
