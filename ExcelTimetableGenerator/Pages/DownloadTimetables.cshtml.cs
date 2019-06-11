using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace ExcelTimetableGenerator.Pages
{
    public class DownloadTimetablesModel : PageModel
    {
        private IHostingEnvironment _hostingEnvironment;

        public DownloadTimetablesModel(
            IHostingEnvironment hostingEnvironment
            )
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public string ZipPath { get; set; }
        public string ZipFileName { get; set; }
        public FileContentResult OnGet()
        {
            ZipPath = _hostingEnvironment.WebRootPath + @"\Output\Timetables.zip";
            ZipFileName = "Timetables.zip";

            byte[] fileBytes = System.IO.File.ReadAllBytes(ZipPath);

            return File(fileBytes, "application/force-download", ZipFileName);
        }
    }
}