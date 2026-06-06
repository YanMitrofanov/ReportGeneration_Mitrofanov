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
    public class WorkContext : Work
    {
        public WorkContext(int Id, int IdDiscipline, int IdType, DateTime Date, string Name, int Semester)
            : base(Id, IdDiscipline, IdType, Date, Name, Semester) { }

        public static List<WorkContext> AllWorks()
        {
            List<WorkContext> allWorks = new List<WorkContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader reader = Connection.Query("SELECT * FROM work ORDER BY Date;", connection);

            while (reader.Read())
            {
                allWorks.Add(new WorkContext(
                    reader.GetInt32(0),
                    reader.GetInt32(1),
                    reader.GetInt32(2),
                    reader.GetDateTime(3),
                    reader.GetString(4),
                    reader.GetInt32(5)
                ));
            }

            Connection.CloseConnection(connection);
            return allWorks;
        }
    }
}
