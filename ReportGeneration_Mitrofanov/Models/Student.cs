using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Mitrofanov.Models
{
    public class Student
    {
        public int Id { get; set; }
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public int IdGroup { get; set; }
        public bool Expelled { get; set; }
        public DateTime DateTimeExpelled { get; set; }

        public Student(int id, string firstname, string lastname, int idGroup, bool expelled, DateTime dateTimeExpelled)
        {
            Id = id;
            Firstname = firstname;
            Lastname = lastname;
            IdGroup = idGroup;
            Expelled = expelled;
            DateTimeExpelled = dateTimeExpelled;
        }
    }
}
