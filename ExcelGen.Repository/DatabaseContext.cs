using ExcelGen.Repository.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelGen.Repository
{
    public class DatabaseContext : DbContext
    {
        public DatabaseContext(DbContextOptions options) : base(options) 
        {

        }
        public DbSet<Discipline> Discipline { get; set; }
        public DbSet<Semester> Semester { get; set; }
    }
}
