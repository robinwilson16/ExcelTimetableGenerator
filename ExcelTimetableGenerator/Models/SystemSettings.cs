using ExcelTimetableGenerator.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class SystemSettings
    {
        public SystemSettings()
        {
            Greeting = Identity.GetGreeting();
        }

        [Display(Name = "System Version Number")]
        public string Greeting { get; set; }

        public string UserDetails { get; set; }
        public string PlanningSystem { get; set; }
        public int MinutesToKeepFolders { get; set; }
        public int ProgrammeNameMaxLength { get; set; }

        public string Version { get; set; }
    }
}
