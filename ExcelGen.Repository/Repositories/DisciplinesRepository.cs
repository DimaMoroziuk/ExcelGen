using ExcelGen.Repository.Interfaces;
using ExcelGen.Repository.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGen.Repository.Repositories
{
    public class DisciplinesRepository : IDisciplineRepository
    {
        private readonly DatabaseContext _context;

        public DisciplinesRepository(DatabaseContext context)
        {
            _context = context;
        }

        public async Task<List<Discipline>> GetAllDisciplines() 
        {
            return await _context.Discipline.ToListAsync();
        }

        public async Task<Discipline> GetDisciplineById(string id)
        {
            var disciplineDTO = await _context.Discipline.FindAsync(id);

            if (disciplineDTO == null)
            {
                throw new Exception();
            }

            return disciplineDTO;
        }

        public async Task UpdateExistingDiscipline(string id, Discipline discipline) 
        {
            if (id != discipline.Id)
            {
                throw new Exception();
            }

            _context.Entry(discipline).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                DisciplineDTOExists(id);

                throw;
            }

            return;
        }

        public async Task CreateNewDiscipline(Discipline discipline) 
        {
            _context.Discipline.Add(discipline);
            await _context.SaveChangesAsync();

            return;
        }

        public async Task DeleteDisciplineById(string id) 
        {
            var discipline = await _context.Discipline.FindAsync(id);
            if (discipline == null)
            {
                throw new Exception();
            }

            _context.Discipline.Remove(discipline);
            await _context.SaveChangesAsync();

            return;

        }

        #region Private Methods

        private void DisciplineDTOExists(string id)
        {
            if (_context.Discipline.All(e => e.Id != id))
                throw new Exception();
        }

        #endregion
    }
}
