using ExcelGen.Repository.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGen.Repository.Interfaces
{
    public interface ISemesterRepository
    {
        Task<List<Semester>> GetSemestersByDisciplineId(string id);
        Task DeleteSemesterByDisciplineId(string id);
        Task CreateNewSemesters(List<Semester> semesters);
        Task CreateNewSemester(Semester semester);
    }
}
