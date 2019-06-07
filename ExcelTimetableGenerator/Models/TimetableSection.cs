using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class TimetableSection
    {
        public int TimetableSectionID { get; set; }

        [StringLength(50)]
        public string SectionName { get; set; }
    }
}
