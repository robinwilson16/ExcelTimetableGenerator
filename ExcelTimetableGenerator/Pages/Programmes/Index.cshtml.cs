using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;

namespace ExcelTimetableGenerator.Pages.Programmes
{
    public class IndexModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;

        public IndexModel(ExcelTimetableGenerator.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        public IList<Programme> Programme { get;set; }

        public async Task OnGetAsync()
        {
            Programme = await _context.Programme.ToListAsync();
        }
    }
}
