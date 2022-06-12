using ExcelGen.Repository.Interfaces;
using ExcelGen.Repository.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGen.Repository.Repositories
{
    public class SemesterRepository : ISemesterRepository
    { 
        private readonly DatabaseContext _context;

        public SemesterRepository(DatabaseContext context)
        {
            _context = context;
        }

        public async Task<List<Semester>> GetSemestersByDisciplineId(string id)
        {
            var semesters = await _context.Semester.Where(sem => sem.DisciplineId == id).ToListAsync();

            if (semesters == null)
            {
                throw new Exception();
            }

            return semesters;
        }

        public async Task UpdateExistingSemesters(IEnumerable<string> ids, IEnumerable<Semester> semesters)
        {
            foreach (var sem in semesters)
            {
                if (ids.Any(id => sem.Id != id))
                {
                    throw new Exception();
                }
                else
                {
                    _context.Entry(sem).State = EntityState.Modified;
                }
            }

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                throw;
            }
        }

        public async Task CreateNewSemester(Semester semester)
        {
            _context.Semester.Add(semester);
            await _context.SaveChangesAsync();

            return;
        }

        public async Task CreateNewSemesters(List<Semester> semesters)
        {
            _context.Semester.AddRange(semesters);
            await _context.SaveChangesAsync();

            return;
        }

        public async Task DeleteSemesterByDisciplineId(string id)
        {
            var semesters = await _context.Semester.Where(sem => sem.DisciplineId == id).ToListAsync();
            if (semesters == null)
            {
                throw new Exception();
            }

            _context.Semester.RemoveRange(semesters);
            await _context.SaveChangesAsync();

            return;

        }
    }
}
