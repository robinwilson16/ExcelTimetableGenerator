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

        public DbSet<BankHoliday> BankHoliday { get; set; }
        public DbSet<Config> Config { get; set; }
        public DbSet<Course> Course { get; set; }
        public DbSet<Day> Day { get; set; }
        public DbSet<Group> Group { get; set; }
        public DbSet<Programme> Programme { get; set; }
        public DbSet<SelectListData> SelectListData { get; set; }
        public DbSet<StaffMember> StaffMember { get; set; }
        public DbSet<TermDate> TermDate { get; set; }
        public DbSet<Time> Time { get; set; }
        public DbSet<TimetableSection> TimetableSection { get; set; }

        public DbSet<Week> Week { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //Needed to add composite key
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<Config>()
                .HasKey(c => new { c.AcademicYear });

            modelBuilder.Entity<SelectListData>()
                .HasKey(d => new { d.Code });

            modelBuilder.Entity<StaffMember>()
                .HasKey(c => new { c.StaffRef });
        }
    }
}
