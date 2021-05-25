using ExcelTimetableGenerator.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(
            DbContextOptions<ApplicationDbContext> options,
            IConfiguration _configuration)
            : base(options)
        {
            configuration = _configuration;
        }

        public IConfiguration configuration { get; }

        public DbSet<BankHoliday> BankHoliday { get; set; }
        public DbSet<Config> Config { get; set; }
        public DbSet<Course> Course { get; set; }
        public DbSet<DaySlot> DaySlot { get; set; }
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

            modelBuilder.Entity<SystemSettings>()
                .HasNoKey();

            //Prevent creating table in EF Migration
            modelBuilder.Entity<Config>(entity => {
                entity.ToView("Config", "dbo");
            });
            modelBuilder.Entity<Course>(entity => {
                entity.ToView("Course", "dbo");
            });
            modelBuilder.Entity<Group>(entity => {
                entity.ToView("Group", "dbo");
            });
            modelBuilder.Entity<Programme>(entity => {
                entity.ToView("Programme", "dbo");
            });
            modelBuilder.Entity<SelectListData>(entity => {
                entity.ToView("SelectListData", "dbo");
            });
            modelBuilder.Entity<StaffMember>(entity => {
                entity.ToView("StaffMember", "dbo");
            });
            modelBuilder.Entity<SystemSettings>(entity => {
                entity.ToView("SystemSettings", "dbo");
            });
            modelBuilder.Entity<SystemSettings>(entity => {
                entity.ToView("SystemSettings", "dbo");
            });
            modelBuilder.Entity<Time>(entity => {
                entity.ToView("Time", "dbo");
            });
        }

        //Rename migration history table
        protected override void OnConfiguring(DbContextOptionsBuilder options)
            => options.UseSqlServer(
                configuration.GetConnectionString("DefaultConnection"),
                x => x.MigrationsHistoryTable("__ETG_EFMigrationsHistory", "dbo"));

        //Rename migration history table
}
}
