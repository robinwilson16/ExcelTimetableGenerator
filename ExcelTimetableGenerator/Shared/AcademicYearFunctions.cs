using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Shared
{
    public class AcademicYearFunctions
    {
        public static Config ConfigData { get; set; }
        public static async Task<string> GetAcademicYear(string academicYear, ApplicationDbContext _context)
        {
            string currentAcademicYear;

            if (academicYear == null)
            {
                //Get value from database
                ConfigData = (await _context.Config
                    .FromSqlInterpolated($"EXEC SPR_ETG_Config")
                    .ToListAsync())
                    .FirstOrDefault();

                currentAcademicYear = ConfigData.AcademicYear;
            }
            else
            {
                currentAcademicYear = academicYear;
            }

            return currentAcademicYear;
        }

        public static async Task<string> GetDefaultAcademicYear(ApplicationDbContext _context)
        {
            //Get value from database
            ConfigData = (await _context.Config
                .FromSqlInterpolated($"EXEC SPR_ETG_Config")
                .ToListAsync())
                .FirstOrDefault();

            return ConfigData.AcademicYear;
        }
    }
}
