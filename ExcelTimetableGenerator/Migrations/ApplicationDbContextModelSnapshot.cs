﻿// <auto-generated />
using System;
using ExcelTimetableGenerator.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace ExcelTimetableGenerator.Migrations
{
    [DbContext(typeof(ApplicationDbContext))]
    partial class ApplicationDbContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .UseIdentityColumns()
                .HasAnnotation("Relational:MaxIdentifierLength", 128)
                .HasAnnotation("ProductVersion", "5.0.2");

            modelBuilder.Entity("ExcelTimetableGenerator.Models.BankHoliday", b =>
                {
                    b.Property<int>("BankHolidayID")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("AcademicYearID")
                        .HasMaxLength(5)
                        .HasColumnType("nvarchar(5)");

                    b.Property<string>("BankHolidayDesc")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.HasKey("BankHolidayID");

                    b.ToTable("ETG_BankHoliday");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Config", b =>
                {
                    b.Property<string>("AcademicYear")
                        .HasMaxLength(5)
                        .HasColumnType("nvarchar(5)");

                    b.HasKey("AcademicYear");

                    b.ToView("Config", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Course", b =>
                {
                    b.Property<int>("CourseID")
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("AimCode")
                        .HasMaxLength(8)
                        .HasColumnType("nvarchar(8)");

                    b.Property<string>("AwardBody")
                        .HasMaxLength(150)
                        .HasColumnType("nvarchar(150)");

                    b.Property<string>("CourseCode")
                        .HasMaxLength(72)
                        .HasColumnType("nvarchar(72)");

                    b.Property<int>("CourseOrder")
                        .HasColumnType("int");

                    b.Property<string>("CourseStatus")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<string>("CourseTitle")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.Property<int?>("EEP1618")
                        .HasColumnType("int");

                    b.Property<int?>("EEP19")
                        .HasColumnType("int");

                    b.Property<int?>("EEPMax")
                        .HasColumnType("int");

                    b.Property<DateTime?>("EndDate")
                        .HasColumnType("datetime2");

                    b.Property<string>("FundModel")
                        .HasMaxLength(20)
                        .HasColumnType("nvarchar(20)");

                    b.Property<string>("FundSource")
                        .HasMaxLength(3)
                        .HasColumnType("nvarchar(3)");

                    b.Property<string>("FundStream")
                        .HasMaxLength(2)
                        .HasColumnType("nvarchar(2)");

                    b.Property<int?>("GroupSize")
                        .HasColumnType("int");

                    b.Property<decimal?>("HoursPerWeek")
                        .HasColumnType("decimal(9,2)");

                    b.Property<bool?>("IsMainCourse")
                        .HasColumnType("bit");

                    b.Property<string>("ModeOfAttendanceCode")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("ModeOfAttendanceName")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<int?>("NoOfYears")
                        .HasColumnType("int");

                    b.Property<string>("Notes")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int?>("NumGroups")
                        .HasColumnType("int");

                    b.Property<int?>("PLH1618")
                        .HasColumnType("int");

                    b.Property<int?>("PLH19")
                        .HasColumnType("int");

                    b.Property<int?>("PLHMax")
                        .HasColumnType("int");

                    b.Property<string>("ProgType")
                        .HasMaxLength(2)
                        .HasColumnType("nvarchar(2)");

                    b.Property<int>("ProgrammeID")
                        .HasColumnType("int");

                    b.Property<string>("SiteCode")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("SiteName")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<DateTime?>("StartDate")
                        .HasColumnType("datetime2");

                    b.Property<int?>("Weeks")
                        .HasColumnType("int");

                    b.Property<int?>("YearNo")
                        .HasColumnType("int");

                    b.HasKey("CourseID");

                    b.HasIndex("ProgrammeID");

                    b.ToView("Course", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.DaySlot", b =>
                {
                    b.Property<int>("DaySlotID")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("AcademicYearID")
                        .HasMaxLength(5)
                        .HasColumnType("nvarchar(5)");

                    b.Property<string>("DayName")
                        .HasMaxLength(20)
                        .HasColumnType("nvarchar(20)");

                    b.Property<int>("NumSlots")
                        .HasColumnType("int");

                    b.HasKey("DaySlotID");

                    b.ToTable("ETG_DaySlot");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Group", b =>
                {
                    b.Property<string>("GroupID")
                        .HasMaxLength(20)
                        .HasColumnType("nvarchar(20)");

                    b.Property<string>("GroupCode")
                        .HasMaxLength(2)
                        .HasColumnType("nvarchar(2)");

                    b.Property<string>("ProgCodeWithGroup")
                        .HasMaxLength(75)
                        .HasColumnType("nvarchar(75)");

                    b.Property<int>("ProgrammeID")
                        .HasColumnType("int");

                    b.HasKey("GroupID");

                    b.HasIndex("ProgrammeID");

                    b.ToView("Group", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Programme", b =>
                {
                    b.Property<int>("ProgrammeID")
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<int?>("EEP1618")
                        .HasColumnType("int");

                    b.Property<int?>("EEP19")
                        .HasColumnType("int");

                    b.Property<int?>("EEPMax")
                        .HasColumnType("int");

                    b.Property<string>("FacCode")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("FacName")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("ModeOfAttendanceCode")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("ModeOfAttendanceName")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<int?>("PLH1618")
                        .HasColumnType("int");

                    b.Property<int?>("PLH19")
                        .HasColumnType("int");

                    b.Property<int?>("PLHMax")
                        .HasColumnType("int");

                    b.Property<string>("ProgCode")
                        .HasMaxLength(72)
                        .HasColumnType("nvarchar(72)");

                    b.Property<string>("ProgStatus")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<string>("ProgTitle")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.Property<string>("SiteCode")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("SiteName")
                        .HasMaxLength(250)
                        .HasColumnType("nvarchar(250)");

                    b.Property<string>("TeamCode")
                        .HasMaxLength(24)
                        .HasColumnType("nvarchar(24)");

                    b.Property<string>("TeamName")
                        .HasMaxLength(150)
                        .HasColumnType("nvarchar(150)");

                    b.HasKey("ProgrammeID");

                    b.ToView("Programme", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.SelectListData", b =>
                {
                    b.Property<string>("Code")
                        .HasMaxLength(10)
                        .HasColumnType("nvarchar(10)");

                    b.Property<string>("Description")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.HasKey("Code");

                    b.ToView("SelectListData", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.StaffMember", b =>
                {
                    b.Property<string>("StaffRef")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<string>("Forename")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<string>("StaffDetails")
                        .HasMaxLength(150)
                        .HasColumnType("nvarchar(150)");

                    b.Property<string>("Surname")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.HasKey("StaffRef");

                    b.ToView("StaffMember", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.SystemSettings", b =>
                {
                    b.Property<string>("Greeting")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("MinutesToKeepFolders")
                        .HasColumnType("int");

                    b.Property<string>("PlanningSystem")
                        .HasColumnType("nvarchar(max)");

                    b.Property<int>("ProgrammeNameMaxLength")
                        .HasColumnType("int");

                    b.Property<string>("UserDetails")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Version")
                        .HasColumnType("nvarchar(max)");

                    b.ToView("SystemSettings", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.TermDate", b =>
                {
                    b.Property<int>("TermDateID")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("AcademicYearID")
                        .HasMaxLength(5)
                        .HasColumnType("nvarchar(5)");

                    b.Property<string>("Dates")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.Property<bool>("IsTerm")
                        .HasColumnType("bit");

                    b.Property<string>("TermDateName")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<int>("TermDateOrder")
                        .HasColumnType("int");

                    b.HasKey("TermDateID");

                    b.ToTable("ETG_TermDate");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Time", b =>
                {
                    b.Property<int>("TimeID")
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<int>("Hours")
                        .HasMaxLength(2)
                        .HasColumnType("int");

                    b.Property<int>("Mins")
                        .HasMaxLength(2)
                        .HasColumnType("int");

                    b.Property<string>("TimeName")
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("TimeID");

                    b.ToView("Time", "dbo");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.TimetableSection", b =>
                {
                    b.Property<int>("TimetableSectionID")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("SectionName")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.HasKey("TimetableSectionID");

                    b.ToTable("ETG_TimetableSection");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Week", b =>
                {
                    b.Property<int>("WeekID")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .UseIdentityColumn();

                    b.Property<string>("AcademicYearID")
                        .HasMaxLength(5)
                        .HasColumnType("nvarchar(5)");

                    b.Property<string>("Notes")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.Property<string>("Notes2")
                        .HasMaxLength(255)
                        .HasColumnType("nvarchar(255)");

                    b.Property<string>("WeekDesc")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<int?>("WeekNum")
                        .HasColumnType("int");

                    b.HasKey("WeekID");

                    b.ToTable("ETG_Week");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Course", b =>
                {
                    b.HasOne("ExcelTimetableGenerator.Models.Programme", "Programme")
                        .WithMany("Course")
                        .HasForeignKey("ProgrammeID")
                        .OnDelete(DeleteBehavior.Cascade)
                        .IsRequired();

                    b.Navigation("Programme");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Group", b =>
                {
                    b.HasOne("ExcelTimetableGenerator.Models.Programme", "Programme")
                        .WithMany("Group")
                        .HasForeignKey("ProgrammeID")
                        .OnDelete(DeleteBehavior.Cascade)
                        .IsRequired();

                    b.Navigation("Programme");
                });

            modelBuilder.Entity("ExcelTimetableGenerator.Models.Programme", b =>
                {
                    b.Navigation("Course");

                    b.Navigation("Group");
                });
#pragma warning restore 612, 618
        }
    }
}
