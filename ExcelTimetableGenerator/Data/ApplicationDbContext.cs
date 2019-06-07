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
        public DbSet<Day> Day { get; set; }
        public DbSet<Group> Group { get; set; }
        public DbSet<Programme> Programme { get; set; }
        public DbSet<Time> Time { get; set; }
        public DbSet<TimetableSection> TimetableSection { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //Needed to add composite key
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<Config>()
                .HasKey(c => new { c.AcademicYear });
        }
    }
}
