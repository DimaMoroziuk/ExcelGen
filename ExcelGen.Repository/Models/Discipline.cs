using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace ExcelGen.Repository.Models
{
    public class Discipline
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public string Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public int ControlType { get; set; }
        public int NumberOfECTS { get; set; }
        public int LectureHours { get; set; }
        public int LabHours { get; set; }
        public int PracticeHours { get; set; }
        public int Category { get; set; }
    }
}
