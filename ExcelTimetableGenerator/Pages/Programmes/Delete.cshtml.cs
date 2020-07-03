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
    public class DeleteModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;

        public DeleteModel(ExcelTimetableGenerator.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        [BindProperty]
        public Programme Programme { get; set; }

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            Programme = await _context.Programme.FirstOrDefaultAsync(m => m.ProgrammeID == id);

            if (Programme == null)
            {
                return NotFound();
            }
            return Page();
        }

        public async Task<IActionResult> OnPostAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            Programme = await _context.Programme.FindAsync(id);

            if (Programme != null)
            {
                _context.Programme.Remove(Programme);
                await _context.SaveChangesAsync();
            }

            return RedirectToPage("./Index");
        }
    }
}
