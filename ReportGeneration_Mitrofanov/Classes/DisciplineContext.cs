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
    public class DisciplineContext : Discipline
    {
        public DisciplineContext(int Id, string Name, int IdGroup) : base(Id, Name, IdGroup) { }

        public static List<DisciplineContext> AllDisciplines()
        {
            List<DisciplineContext> allDisciplines = new List<DisciplineContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader reader = Connection.Query("SELECT * FROM discipline ORDER BY Name;", connection);

            while (reader.Read())
            {
                allDisciplines.Add(new DisciplineContext(
                    reader.GetInt32(0),
                    reader.GetString(1),
                    reader.GetInt32(2)
                ));
            }

            Connection.CloseConnection(connection);
            return allDisciplines;
        }
    }
}
