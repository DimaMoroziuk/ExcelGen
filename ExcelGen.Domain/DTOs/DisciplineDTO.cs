using ExcelGen.Domain.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace ExcelGen.Domain.DTOs
{
    public class DisciplineDTO
    {
        private eControlType _controlType;
        private eDisciplineCategory _disciplineCategory;
        public string Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public eControlType ControlType { get { return _controlType; } set
            {
                _controlType = value;
                switch (value) 
                {
                    case eControlType.Test:
                        UkrainianControlType = "Залік";
                        break;
                    case eControlType.Examination:
                        UkrainianControlType = "Екзамен";
                        break;
                    case eControlType.CourseProject:
                        UkrainianControlType = "Курсова робота";
                        break;
                    default:
                        UkrainianControlType = "Помилковий тип контролю";
                        break;
                }
                 
                    } }
        public int NumberOfECTS { get; set; }
        public int LectureHours { get; set; }
        public int LabHours { get; set; }
        public int PracticeHours { get; set; }
        public eDisciplineCategory Category
        {
            get { return _disciplineCategory; }
            set
            {
                _disciplineCategory = value;
                switch (value)
                {
                    case eDisciplineCategory.Necessary:
                        UkrainianCategory = "Обов'язкова";
                        break;
                    case eDisciplineCategory.Chosen:
                        UkrainianCategory = "Вибіркова";
                        break;
                    default:
                        UkrainianCategory = "Помилкова котегорія";
                        break;
                }

            }
        }
        public Dictionary<int, int> SemestersHours { get; set; }
        public string UkrainianControlType { get; set; }
        public string UkrainianCategory { get; set; }
    }
}
