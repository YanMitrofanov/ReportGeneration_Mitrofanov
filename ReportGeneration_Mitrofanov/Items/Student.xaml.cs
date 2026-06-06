using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReportGeneration_Mitrofanov.Models;
using ReportGeneration_Mitrofanov.Classes;

namespace ReportGeneration_Mitrofanov.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public int StudentId { get; set; }
        public int IdGroup { get; set; }
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public Student(int id, string firstname, string lastname, int idGroup, bool expelled)
        {
            InitializeComponent();
            StudentId = id;
            Firstname = firstname;
            Lastname = lastname;
            IdGroup = idGroup;

            TBFio.Text = $"{lastname} {firstname}";
            CBExemplel.IsChecked = expelled;

            LoadStudentProgress();
        }
        private void LoadStudentProgress()
        {
            
            var allEvaluations = EvaluationContext.AllEvaluations();
            var allWorks = WorkContext.AllWorks();

            var studentEvals = allEvaluations.Where(e => e.IdStudent == StudentId).ToList();
            var totalWorks = allWorks.Count();

            if (totalWorks > 0)
            {
                double completed = studentEvals.Count(e => e.Value == "зачет" || e.Value == "отлично" || e.Value == "хорошо" || e.Value == "удовлетворительно");
                doneWorks.Value = (completed / totalWorks) * 100;
            }

            
            int latenessCount = studentEvals.Count(e => int.TryParse(e.Lateness, out int lateness) && lateness > 0);
            missedCount.Value = latenessCount;
        }
    }
}
