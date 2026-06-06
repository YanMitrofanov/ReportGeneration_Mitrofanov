using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ReportGeneration_Mitrofanov.Pages;
using System.Windows;

namespace ReportGeneration_Mitrofanov.Classes.Common
{
    public class Report
    {
        /// <summary> Метод создания отчёта о группе </summary>
        public static void GenerateGroupReport(int IdGroup, Main MainPage)
        {
            // Создаём диалог для сохранения
            SaveFileDialog SFD = new SaveFileDialog
            {
                // Указываем начальную директорию
                InitialDirectory = @"C:\",
                // Указываем формат сохранения файла
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"Отчет_по_группе_{IdGroup}_{System.DateTime.Now:yyyyMMdd_HHmmss}"
            };

            // Открываем диалоговое окно
            if (SFD.ShowDialog() == true)
            {
                // Получаем группу, о которой сохраняем информацию
                var Group = MainPage.AllGroups.Find(x => x.Id == IdGroup);
                if (Group == null) return;

                // Открываем Excel
                var ExcelApp = new Excel.Application();
                try
                {
                    // Скрываем его видимость
                    ExcelApp.Visible = false;
                    // Добавляем книгу
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    // Получаем активный лист
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    // Обращаемся к ячейке A1 и указываем текст
                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    // Объединяем ячейки A1 и E1
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    // Применяем стили
                    ApplyStyles(Worksheet.Cells[1, 1], 18);

                    // Заголовки таблицы
                    (Worksheet.Cells[3, 1] as Excel.Range).Value = "ФИО студента";
                    (Worksheet.Cells[3, 2] as Excel.Range).Value = "Выполненные работы (%)";
                    (Worksheet.Cells[3, 3] as Excel.Range).Value = "Количество опозданий";
                    (Worksheet.Cells[3, 4] as Excel.Range).Value = "Статус";

                    // Применяем стили к заголовкам
                    for (int col = 1; col <= 4; col++)
                    {
                        ApplyStyles(Worksheet.Cells[3, col], 12, true);
                    }

                    // Получаем студентов группы
                    var students = MainPage.AllStudents.Where(s => s.IdGroup == IdGroup && !s.Expelled).ToList();

                    int row = 4;
                    foreach (var student in students)
                    {
                        // Получаем оценки студента
                        var evaluations = MainPage.AllEvaluations.Where(e => e.IdStudent == student.Id).ToList();

                        // Все работы по дисциплинам группы
                        var groupDisciplineIds = MainPage.AllDisciplines.Where(d => d.IdGroup == IdGroup).Select(d => d.Id);
                        var groupWorks = MainPage.AllWorks.Where(w => groupDisciplineIds.Contains(w.IdDiscipline)).ToList();

                        int totalWorks = groupWorks.Count;
                        int completedWorks = 0;
                        int latenessCount = 0;

                        foreach (var eval in evaluations)
                        {
                            var work = MainPage.AllWorks.FirstOrDefault(w => w.Id == eval.IdWork);
                            if (work != null && groupDisciplineIds.Contains(work.IdDiscipline))
                            {
                                if (eval.Value == "зачет" || eval.Value == "отлично" ||
                                    eval.Value == "хорошо" || eval.Value == "удовлетворительно" ||
                                    eval.Value == "5" || eval.Value == "4" || eval.Value == "3")
                                {
                                    completedWorks++;
                                }

                                if (int.TryParse(eval.Lateness, out int lateness))
                                {
                                    latenessCount += lateness;
                                }
                            }
                        }

                        double percentCompleted = totalWorks > 0 ? (double)completedWorks / totalWorks * 100 : 0;

                        (Worksheet.Cells[row, 1] as Excel.Range).Value = $"{student.Lastname} {student.Firstname}";
                        (Worksheet.Cells[row, 2] as Excel.Range).Value = $"{percentCompleted:F1}%";
                        (Worksheet.Cells[row, 3] as Excel.Range).Value = latenessCount;
                        (Worksheet.Cells[row, 4] as Excel.Range).Value = student.Expelled ? "Отчислен" : "Учится";

                        row++;
                    }

                    // Сохраняем файл
                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();

                    MessageBox.Show($"Отчёт успешно сохранён!\n{SFD.FileName}", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                finally
                {
                    ExcelApp.Quit();
                }
            }
        }

        private static void ApplyStyles(Excel.Range cell, int fontSize, bool isBold = false)
        {
            cell.Font.Size = fontSize;
            cell.Font.Bold = isBold;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }
    }
}
