using ExcelTimetableGenerator.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<Config> Config { get; set; }
        public DbSet<Course> Course { get; set; }
        public DbSet<Programme> Programme { get; set; }
        public DbSet<Timetable> Timetable { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //Needed to add composite key
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<Config>()
                .HasKey(c => new { c.AcademicYear });

            modelBuilder.Entity<CustomerArea>()
                .HasKey(c => new { c.CustomerID, c.AreaID });
        }
    }
}
