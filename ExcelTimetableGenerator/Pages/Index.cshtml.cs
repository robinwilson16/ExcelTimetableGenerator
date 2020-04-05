using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;
using ExcelTimetableGenerator.Shared;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace ExcelTimetableGenerator.Pages
{
    [Authorize(Roles = "ALLSTAFF")]
    public class IndexModel : PageModel
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration;

        public IndexModel(ExcelTimetableGenerator.Data.ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

        public string AcademicYear { get; set; }
        public string UserDetails { get; set; }
        public string UserGreeting { get; set; }
        public string SystemVersion { get; set; }

        public Guid SessionID { get; set; }

        public IList<SelectListData> ProgrammeSelectList { get; set; }

        public async Task OnGetAsync(string academicYear, int plan, string course)
        {
            Guid sessionID = Guid.NewGuid();
            //HttpContext.Session.SetString("SessionID", sessionID.ToString());
            SessionID = sessionID;
            //SessionID = Guid.Parse(HttpContext.Session.GetString("SessionID"));

            int planRevisionID = 0;

            if (plan >= 1)
            {
                planRevisionID = plan;
            }
            else
            {
                planRevisionID = int.Parse(_configuration.GetSection("ProResource")["PlanRevisionID"]);
            }

            AcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            UserDetails = await Identity.GetFullName(academicYear, User.Identity.Name.Split('\\').Last(), _context);

            UserGreeting = Identity.GetGreeting();

            SystemVersion = _configuration["Version"];

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            ProgrammeSelectList = await _context.SelectListData
                .FromSqlInterpolated($"EXEC SPR_ETG_ProgrammeSelectList @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}")
                .ToListAsync();

            ViewData["CourseSL"] = new SelectList(ProgrammeSelectList, "Code", "Description");
        }
    }
}
