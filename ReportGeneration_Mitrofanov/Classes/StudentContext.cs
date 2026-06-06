using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGeneration_Mitrofanov.Models;
using ReportGeneration_Mitrofanov.Classes.Common;
using MySql.Data.MySqlClient;

namespace ReportGeneration_Mitrofanov.Classes
{
    public class StudentContext : Student
    {
        public StudentContext(int Id, string Firstname, string Lastname, int IdGroup, bool Expelled, DateTime DateTimeExpelled)
                   : base(Id, Firstname, Lastname, IdGroup, Expelled, DateTimeExpelled) { }

        public static List<StudentContext> AllStudents()
        {
            List<StudentContext> allStudents = new List<StudentContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader reader = Connection.Query("SELECT * FROM student ORDER BY Lastname;", connection);

            while (reader.Read())
            {
                allStudents.Add(new StudentContext(
                    reader.GetInt32(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    reader.GetInt32(3),
                    reader.GetBoolean(4),
                    reader.IsDBNull(5) ? DateTime.Now : reader.GetDateTime(5)
                ));
            }

            Connection.CloseConnection(connection);
            return allStudents;
        }
    }
}
