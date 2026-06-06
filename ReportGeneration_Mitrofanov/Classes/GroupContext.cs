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
    public class GroupContext : Group
    {
        public GroupContext(int Id, string Name) : base(Id, Name) { }

        public static List<GroupContext> AllGroups()
        {
            List<GroupContext> allGroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader reader = Connection.Query("SELECT * FROM `group` ORDER BY Name;", connection);

            while (reader.Read())
            {
                allGroups.Add(new GroupContext(
                    reader.GetInt32(0),
                    reader.GetString(1)
                ));
            }

            Connection.CloseConnection(connection);
            return allGroups;
        }
    }
}
