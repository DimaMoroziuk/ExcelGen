using ExcelGen.Domain.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGen.Domain.Interfaces
{
    public interface IDisciplineManager
    {
        Task<List<DisciplineDTO>> GetAllDisciplines();
        Task<DisciplineDTO> GetDisciplineById(string id);
        Task UpdateExistingDiscipline(string id, DisciplineDTO discipline);
        Task CreateNewDiscipline(DisciplineDTO discipline);
        Task DeleteDisciplineById(string id);
        Task FillWorkbook(ExcelWorkbook workBook);
    }
}
