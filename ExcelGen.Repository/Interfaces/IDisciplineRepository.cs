using ExcelGen.Repository.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGen.Repository.Interfaces
{
    public interface IDisciplineRepository
    {
        Task<List<Discipline>> GetAllDisciplines();
        Task<Discipline> GetDisciplineById(string id);
        Task UpdateExistingDiscipline(string id, Discipline discipline);
        Task CreateNewDiscipline(Discipline discipline);
        Task DeleteDisciplineById(string id);
    }
}
