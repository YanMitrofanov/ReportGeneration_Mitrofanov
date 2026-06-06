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
using ReportGeneration_Mitrofanov.Classes;
using ReportGeneration_Mitrofanov.Items;
using ReportGeneration_Mitrofanov.Classes.Common;
using Mysqlx.Crud;

namespace ReportGeneration_Mitrofanov.Pages
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class Main : Page
    {
        
        public List<GroupContext> AllGroups = GroupContext.AllGroups();
       
        public List<StudentContext> AllStudents = StudentContext.AllStudents();
        
        public List<WorkContext> AllWorks = WorkContext.AllWorks();
      
        public List<EvaluationContext> AllEvaluations = EvaluationContext.AllEvaluations();
        
        public List<DisciplineContext> AllDisciplines = DisciplineContext.AllDisciplines();
        public Main()
        {
            InitializeComponent();
            CreateGroupUI();
            CreateStudents();
        }

        public void CreateGroupUI()
        {
            foreach (GroupContext Group in AllGroups)
                CBGroups.Items.Add(Group.Name);
            CBGroups.Items.Add("Выберите группу");
            CBGroups.SelectedIndex = CBGroups.Items.Count - 1;
        }
        public void CreateStudents()
        {
            Parent.Children.Clear();
            foreach (var student in AllStudents)
            {
                if (!student.Expelled)
                {
                    Parent.Children.Add(new Student(
                        student.Id,
                        student.Firstname,
                        student.Lastname,
                        student.IdGroup,
                        student.Expelled
                    ));
                }
            }
        }
        public void SelectGroup(object sender, SelectionChangedEventArgs e)
        {
            if (CBGroups.SelectedIndex >= 0 && CBGroups.SelectedIndex < AllGroups.Count)
            {
                var selectedGroup = AllGroups[CBGroups.SelectedIndex];
                Parent.Children.Clear();

                foreach (var student in AllStudents.Where(s => s.IdGroup == selectedGroup.Id && !s.Expelled))
                {
                    Parent.Children.Add(new Student(
                        student.Id,
                        student.Firstname,
                        student.Lastname,
                        student.IdGroup,
                        student.Expelled
                    ));
                }
            }
            else if (CBGroups.SelectedIndex == AllGroups.Count)
            {
                CreateStudents();
            }
        }
        public void SelectStudents(object sender, KeyEventArgs e)
        {
            string searchText = TBFIO.Text.ToLower();

            if (string.IsNullOrWhiteSpace(searchText))
            {
                CreateStudents();
                return;
            }

            Parent.Children.Clear();

            var filteredStudents = AllStudents.Where(s =>
                (s.Lastname.ToLower().Contains(searchText) ||
                 s.Firstname.ToLower().Contains(searchText)) &&
                !s.Expelled);

            if (CBGroups.SelectedIndex >= 0 && CBGroups.SelectedIndex < AllGroups.Count)
            {
                var selectedGroup = AllGroups[CBGroups.SelectedIndex];
                filteredStudents = filteredStudents.Where(s => s.IdGroup == selectedGroup.Id);
            }

            foreach (var student in filteredStudents)
            {
                Parent.Children.Add(new Student(
                    student.Id,
                    student.Firstname,
                    student.Lastname,
                    student.IdGroup,
                    student.Expelled
                ));
            }
        }
        /// <summary> Генерация Excel документа </summary>
        public void ReportGeneration(object sender, RoutedEventArgs e)
        {
            if (CBGroups.SelectedIndex >= 0 && CBGroups.SelectedIndex < AllGroups.Count)
            {
                var selectedGroup = AllGroups[CBGroups.SelectedIndex];
                Report.GenerateGroupReport(selectedGroup.Id, this);
            }
            else
            {
                MessageBox.Show("Выберите группу для формирования отчёта!", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
