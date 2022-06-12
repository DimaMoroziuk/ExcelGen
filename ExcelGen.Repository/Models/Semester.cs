using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace ExcelGen.Repository.Models
{
    public class Semester
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public string Id { get; set; }
        public int Number { get; set; }
        public int Credits { get; set; }

        [ForeignKey("DisciplineId")]
        public string DisciplineId { get; set; }
        public Discipline Discipline { get; set; } 
    }
}
