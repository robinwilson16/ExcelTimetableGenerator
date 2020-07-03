using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;
using ExcelTimetableGenerator.Shared;
using Microsoft.Extensions.Configuration;

namespace ExcelTimetableGenerator.Pages.Programmes
{
    public class DetailsModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;
        private readonly IConfiguration _configuration;

        public DetailsModel(ExcelTimetableGenerator.Data.ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

        public Programme Programme { get; set; }

        public async Task<IActionResult> OnGetAsync(string academicYear, int plan, string course)
        {
            if (course == null)
            {
                return NotFound();
            }

            int planRevisionID = 0;

            if (plan >= 1)
            {
                planRevisionID = plan;
            }
            else
            {
                planRevisionID = int.Parse(_configuration.GetSection("ProResource")["PlanRevisionID"]);
            }

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            Programme = (await _context.Programme
                .FromSqlInterpolated($"EXEC SPR_ETG_ProgrammeData @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}, @Course={course}")
                .ToListAsync())
                .FirstOrDefault();

            if (Programme == null)
            {
                return NotFound();
            }
            return Page();
        }

        public async Task<IActionResult> OnGetJsonAsync(string academicYear, int plan, string course)
        {
            int planRevisionID = 0;

            if (plan >= 1)
            {
                planRevisionID = plan;
            }
            else
            {
                planRevisionID = int.Parse(_configuration.GetSection("ProResource")["PlanRevisionID"]);
            }

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            Programme = (await _context.Programme
                .FromSqlInterpolated($"EXEC SPR_ETG_ProgrammeData @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}, @Course={course}")
                .ToListAsync())
                .FirstOrDefault();

            var collectionWrapper = new
            {
                programme = Programme
            };

            return new JsonResult(collectionWrapper);
        }
    }
}
