using ExcelGen.Domain.DTOs;
using ExcelGen.Domain.Enums;
using ExcelGen.Domain.Interfaces;
using ExcelGen.Repository.Interfaces;
using ExcelGen.Repository.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;

namespace ExcelGen.Domain.Managers
{
    public class DisciplinesManager : IDisciplineManager
    {
        private IDisciplineRepository _disciplineRepository;
        private ISemesterRepository _semesterRepository;

        public DisciplinesManager(IDisciplineRepository disciplineRepository, ISemesterRepository semesterRepository)
        {
            _disciplineRepository = disciplineRepository;
            _semesterRepository = semesterRepository;
        }
        public async Task<List<DisciplineDTO>> GetAllDisciplines()
        {
            var disciplineDTOs = new List<DisciplineDTO>();
            var disciplines = await _disciplineRepository.GetAllDisciplines();

            disciplines.ForEach(disc => disciplineDTOs.Add(MapDisciplineToDTO(disc)));
            foreach (var discDTO in disciplineDTOs) 
            {
                SetSemestersForDiscipline(discDTO, await _semesterRepository.GetSemestersByDisciplineId(discDTO.Id));
            }

            return disciplineDTOs;
        }

        public async Task<DisciplineDTO> GetDisciplineById(string id)
        {
            var discipline = await _disciplineRepository.GetDisciplineById(id);
            var relatedSemesters = await _semesterRepository.GetSemestersByDisciplineId(id);
            var disciplineDTO = MapDisciplineToDTO(discipline);

            SetSemestersForDiscipline(disciplineDTO, relatedSemesters);

            return disciplineDTO;
        }

        public async Task UpdateExistingDiscipline(string id, DisciplineDTO disciplineDTO)
        {
            Validate(disciplineDTO);
            var discipline = MapDTOToDiscipline(disciplineDTO);
            var semesters = new List<Semester>();

            await _disciplineRepository.UpdateExistingDiscipline(id, discipline);

            foreach (var semester in disciplineDTO.SemestersHours)
                semesters.Add(new Semester() { Number = semester.Key, Credits = semester.Value, DisciplineId = disciplineDTO.Id});

            await _semesterRepository.DeleteSemesterByDisciplineId(disciplineDTO.Id);
            await _semesterRepository.CreateNewSemesters(semesters);
        }

        public async Task CreateNewDiscipline(DisciplineDTO disciplineDTO)
        {
            Validate(disciplineDTO);
            var discipline = MapDTOToDiscipline(disciplineDTO);
            var semesters = new List<Semester>();

            foreach (var semester in disciplineDTO.SemestersHours)
                semesters.Add(new Semester() { Number = semester.Key, Credits = semester.Value, DisciplineId = disciplineDTO.Id });

            await _disciplineRepository.CreateNewDiscipline(discipline);
            await _semesterRepository.CreateNewSemesters(semesters);
        }

        public async Task DeleteDisciplineById(string id)
        {
            await _disciplineRepository.DeleteDisciplineById(id);
            await _semesterRepository.DeleteSemesterByDisciplineId(id);
        }

        public async Task FillWorkbook(ExcelWorkbook workBook)
        {
            var discpilines = await GetAllDisciplines();

            ExcelWorksheet worksheet = workBook.Worksheets.Add("Дисципліни");

            PopulateHeaders(worksheet);
            var finalIndex = PopulateData(worksheet, discpilines);
            PopulateSummary(worksheet, finalIndex);
            PrepareExcel(worksheet, finalIndex);
        }

        #region Private
        private void PopulateHeaders(ExcelWorksheet worksheet) 
        {
            worksheet.Cells["A1:A2"].Merge = true;
            worksheet.Cells["B1:B2"].Merge = true;
            worksheet.Cells["C1:C2"].Merge = true;
            worksheet.Cells["D1:D2"].Merge = true;
            worksheet.Cells["E1:E2"].Merge = true;
            worksheet.Cells["F1:F2"].Merge = true;
            worksheet.Cells["G1:G2"].Merge = true;
            worksheet.Cells["H1:H2"].Merge = true;
            worksheet.Cells["I1:N1"].Merge = true;
            worksheet.Cells["O1:O2"].Merge = true;

            worksheet.Cells["A1"].Value = "Шифр за ОПП";
            worksheet.Cells["A1"].Style.TextRotation = 90;
            worksheet.Cells["B1"].Value = "Назва освітнього компоненту";
            worksheet.Cells["C1"].Value = "Тип контролю";
            worksheet.Cells["C1"].Style.TextRotation = 90;
            worksheet.Cells["D1"].Value = "Кількість кредитів ECTS";
            worksheet.Cells["D1"].Style.TextRotation = 90;
            worksheet.Cells["E1"].Value = "Загальна кількість годин";
            worksheet.Cells["E1"].Style.TextRotation = 90;
            worksheet.Cells["F1"].Value = "Лекційні години";
            worksheet.Cells["F1"].Style.TextRotation = 90;
            worksheet.Cells["G1"].Value = "Лабораторні години";
            worksheet.Cells["G1"].Style.TextRotation = 90;
            worksheet.Cells["H1"].Value = "Практичні години";
            worksheet.Cells["H1"].Style.TextRotation = 90;
            worksheet.Cells["I1"].Value = "Розподіл кредитів ECTS за семестрами";
            worksheet.Cells["O1"].Value = "Категорія";

            worksheet.Cells["I2"].Value = "1 Семестр";
            worksheet.Cells["J2"].Value = "2 Семестр";
            worksheet.Cells["K2"].Value = "3 Семестр";
            worksheet.Cells["L2"].Value = "4 Семестр";
            worksheet.Cells["M2"].Value = "5 Семестр";
            worksheet.Cells["N2"].Value = "6 Семестр";

            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(2).Style.Font.Bold = true;
        }

        private int PopulateData(ExcelWorksheet worksheet, List<DisciplineDTO> disciplineDTOs)
        {
            int firstRow = 3;
            var (averageGenHours, averageLabHours, averageLectureHours, averagePracticeHours) = CalculateAverages(disciplineDTOs);
            foreach (var discipline in disciplineDTOs)
            {
                worksheet.Cells[$"A{firstRow}"].Value = discipline.Code;
                worksheet.Cells[$"B{firstRow}"].Value = discipline.Name;
                worksheet.Cells[$"C{firstRow}"].Value = discipline.UkrainianControlType;
                worksheet.Cells[$"D{firstRow}"].Value = discipline.NumberOfECTS;
                worksheet.Cells[$"E{firstRow}"].Value = discipline.LabHours + discipline.LectureHours + discipline.PracticeHours;
                TryMarkAsHigh(worksheet.Cells[$"E{firstRow}"], discipline.LabHours + discipline.LectureHours + discipline.PracticeHours, averageGenHours);
                worksheet.Cells[$"F{firstRow}"].Value = discipline.LectureHours;
                TryMarkAsHigh(worksheet.Cells[$"F{firstRow}"], discipline.LectureHours, averageLectureHours);
                worksheet.Cells[$"G{firstRow}"].Value = discipline.LabHours;
                TryMarkAsHigh(worksheet.Cells[$"G{firstRow}"], discipline.LabHours, averageLabHours);
                worksheet.Cells[$"H{firstRow}"].Value = discipline.PracticeHours;
                TryMarkAsHigh(worksheet.Cells[$"H{firstRow}"], discipline.PracticeHours, averagePracticeHours);
                foreach (var semester in discipline.SemestersHours)
                {
                    FillSemesterHours(worksheet, semester.Key, semester.Value, firstRow);
                }
                worksheet.Cells[$"O{firstRow}"].Value = discipline.UkrainianCategory;
                firstRow++;
            }
            return firstRow;
        }

        private void TryMarkAsHigh(ExcelRange cells, int value, int average)
        {
            if (value > average * 1.5)
            {
                cells.Style.Fill.SetBackground(Color.Red);
                return;
            }
            if (value > average)
            {
                cells.Style.Fill.SetBackground(Color.Yellow);
            }
        }

        private (int, int, int, int) CalculateAverages(List<DisciplineDTO> disciplineDTOs)
        {
            int averageGenHours = 0;
            int averageLabHours = 0;
            int averageLectureHours = 0;
            int averagePracticeHours = 0;
            disciplineDTOs.ForEach(disc => 
                {
                    averageGenHours += (disc.LabHours + disc.LectureHours + disc.PracticeHours) / disciplineDTOs.Count;
                    averageLabHours += disc.LabHours / disciplineDTOs.Count;
                    averageLectureHours += disc.LectureHours / disciplineDTOs.Count;
                    averagePracticeHours += disc.PracticeHours / disciplineDTOs.Count;
                });
            return (averageGenHours, averageLabHours, averageLectureHours, averagePracticeHours);
        }

        // another way of populating summary, where all calculations done on backend
        private void PopulateSummary(ExcelWorksheet worksheet, List<DisciplineDTO> disciplineDTOs, int finalRowIndex)
        {
            worksheet.Cells[$"A{finalRowIndex}:C{finalRowIndex}"].Merge = true;
            worksheet.Cells[$"A{finalRowIndex}"].Value = "Всього : ";

            int sumECTS = 0;
            int sumGenHours = 0;
            int sumLabHours = 0;
            int sumLecHours = 0;
            int sumPracHours = 0;
            disciplineDTOs.ForEach(disc =>
                {
                    sumECTS += disc.NumberOfECTS;
                    sumGenHours += disc.LabHours + disc.LectureHours + disc.PracticeHours;
                    sumLabHours += disc.LabHours;
                    sumLecHours += disc.LectureHours;
                    sumPracHours += disc.PracticeHours;
                });

            worksheet.Cells[$"D{finalRowIndex}"].Value = sumECTS;
            worksheet.Cells[$"E{finalRowIndex}"].Value = sumGenHours;
            worksheet.Cells[$"F{finalRowIndex}"].Value = sumLecHours;
            worksheet.Cells[$"G{finalRowIndex}"].Value = sumLabHours;
            worksheet.Cells[$"H{finalRowIndex}"].Value = sumPracHours;
            for (int i = 1; i < 7; i++)
            {
                int sumForSemester = 0;
                foreach (var disc in disciplineDTOs)
                {
                    foreach (var sem in disc.SemestersHours)
                    {
                        if (sem.Key == i)
                            sumForSemester += sem.Value;
                    }
                }
                FillSemesterHours(worksheet, i, sumForSemester, finalRowIndex);
            }

            worksheet.Row(finalRowIndex).Style.Font.Bold = true;
        }

        private void PopulateSummary(ExcelWorksheet worksheet, int finalRowIndex)
        {
            worksheet.Cells[$"A{finalRowIndex}:C{finalRowIndex}"].Merge = true;
            worksheet.Cells[$"A{finalRowIndex}"].Value = "Всього : ";

            worksheet.Calculate();

            worksheet.Cells[$"D{finalRowIndex}"].Formula = $"SUM(D3:D{finalRowIndex - 1})";
            worksheet.Cells[$"E{finalRowIndex}"].Formula = $"SUM(E3:E{finalRowIndex - 1})";
            worksheet.Cells[$"F{finalRowIndex}"].Formula = $"SUM(F3:F{finalRowIndex - 1})";
            worksheet.Cells[$"G{finalRowIndex}"].Formula = $"SUM(G3:G{finalRowIndex - 1})";
            worksheet.Cells[$"H{finalRowIndex}"].Formula = $"SUM(H3:H{finalRowIndex - 1})";
            worksheet.Cells[$"D{finalRowIndex}"].Formula = $"SUM(D3:D{finalRowIndex - 1})";
            worksheet.Cells[$"E{finalRowIndex}"].Formula = $"SUM(E3:E{finalRowIndex - 1})";
            worksheet.Cells[$"F{finalRowIndex}"].Formula = $"SUM(F3:F{finalRowIndex - 1})";
            worksheet.Cells[$"G{finalRowIndex}"].Formula = $"SUM(G3:G{finalRowIndex - 1})";
            worksheet.Cells[$"H{finalRowIndex}"].Formula = $"SUM(H3:H{finalRowIndex - 1})";
            worksheet.Cells[$"I{finalRowIndex}"].Formula = $"SUM(I3:I{finalRowIndex - 1})";
            worksheet.Cells[$"J{finalRowIndex}"].Formula = $"SUM(J3:J{finalRowIndex - 1})";
            worksheet.Cells[$"K{finalRowIndex}"].Formula = $"SUM(K3:K{finalRowIndex - 1})";
            worksheet.Cells[$"L{finalRowIndex}"].Formula = $"SUM(L3:L{finalRowIndex - 1})";
            worksheet.Cells[$"M{finalRowIndex}"].Formula = $"SUM(M3:M{finalRowIndex - 1})";
            worksheet.Cells[$"N{finalRowIndex}"].Formula = $"SUM(N3:N{finalRowIndex - 1})";
            worksheet.Row(finalRowIndex).Style.Font.Bold = true;

            worksheet.Calculate();
        }

        private void PrepareExcel(ExcelWorksheet worksheet, int finalRowIndex)
        {
            worksheet.Cells[$"A1:O{finalRowIndex}"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheet.Row(2).Height = 150;
            worksheet.Column(2).Width = 40;
            worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        private void FillSemesterHours(ExcelWorksheet worksheet, int semester, int hours, int rowNumber) 
        {
            switch (semester)
            {
                case 1:
                    worksheet.Cells[$"I{rowNumber}"].Value = hours;
                    break;
                case 2:
                    worksheet.Cells[$"J{rowNumber}"].Value = hours;
                    break;
                case 3:
                    worksheet.Cells[$"K{rowNumber}"].Value = hours;
                    break;
                case 4:
                    worksheet.Cells[$"L{rowNumber}"].Value = hours;
                    break;
                case 5:
                    worksheet.Cells[$"M{rowNumber}"].Value = hours;
                    break;
                case 6:
                    worksheet.Cells[$"N{rowNumber}"].Value = hours;
                    break;
                default:
                    break;
            }
        }
        private void Validate(DisciplineDTO disciplineDTO)
        {
            var sumOfECTS = 0;
            foreach (var semester in disciplineDTO.SemestersHours)
                sumOfECTS += semester.Value;

            if (disciplineDTO.NumberOfECTS != sumOfECTS)
                throw new Exception();
        }

        private Discipline MapDTOToDiscipline(DisciplineDTO disciplineDTO)
        {
            var discipline = new Discipline()
            {
                Id = disciplineDTO.Id,
                Code = disciplineDTO.Code,
                Name = disciplineDTO.Name,
                ControlType = (byte)disciplineDTO.ControlType,
                NumberOfECTS = disciplineDTO.NumberOfECTS,
                LectureHours = disciplineDTO.LectureHours,
                LabHours = disciplineDTO.LabHours,
                PracticeHours = disciplineDTO.PracticeHours,
                Category = (byte)disciplineDTO.Category
            };

            return discipline;
        }

        private DisciplineDTO MapDisciplineToDTO(Discipline discipline)
        {
            var disciplineDTO = new DisciplineDTO()
            {
                Id = discipline.Id,
                Code = discipline.Code,
                Name = discipline.Name,
                ControlType = (eControlType)discipline.ControlType,
                NumberOfECTS = discipline.NumberOfECTS,
                LectureHours = discipline.LectureHours,
                LabHours = discipline.LabHours,
                PracticeHours = discipline.PracticeHours,
                Category = (eDisciplineCategory)discipline.Category,
                SemestersHours = new Dictionary<int, int>()
            };

            return disciplineDTO;
        }


        private void SetSemestersForDiscipline (DisciplineDTO disciplineDTO, List<Semester> semesters) 
        {
            if (semesters.Count > 0)
            {
                semesters.ForEach( sem => disciplineDTO.SemestersHours.Add(sem.Number, sem.Credits));
            }
            else 
            {
                return;
            }
        }
        #endregion
    }
}
