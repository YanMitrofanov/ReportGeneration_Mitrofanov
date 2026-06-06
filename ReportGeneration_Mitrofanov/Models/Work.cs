using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Mitrofanov.Models
{
    public class Work
    {
        public int Id { get; set; }
        public int IdDiscipline { get; set; }
        public int IdType { get; set; }
        public DateTime DateTimeDate { get; set; }
        public string Name { get; set; }
        public int Semester { get; set; }

        public Work(int id, int idDiscipline, int idType, DateTime date, string name, int semester)
        {
            Id = id;
            IdDiscipline = idDiscipline;
            IdType = idType;
            DateTimeDate = date;
            Name = name;
            Semester = semester;
        }
    }
}
