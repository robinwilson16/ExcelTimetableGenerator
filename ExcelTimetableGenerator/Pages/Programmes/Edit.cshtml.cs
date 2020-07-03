using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;

namespace ExcelTimetableGenerator.Pages.Programmes
{
    public class EditModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;

        public EditModel(ExcelTimetableGenerator.Data.ApplicationDbContext context)
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

        // To protect from overposting attacks, enable the specific properties you want to bind to, for
        // more details, see https://aka.ms/RazorPagesCRUD.
        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            _context.Attach(Programme).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ProgrammeExists(Programme.ProgrammeID))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return RedirectToPage("./Index");
        }

        private bool ProgrammeExists(int id)
        {
            return _context.Programme.Any(e => e.ProgrammeID == id);
        }
    }
}
