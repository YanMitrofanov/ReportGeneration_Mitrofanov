using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Mitrofanov.Models
{
    public class Discipline
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int IdGroup { get; set; }

        public Discipline(int id, string name, int idGroup)
        {
            Id = id;
            Name = name;
            IdGroup = idGroup;
        }
    }
}
