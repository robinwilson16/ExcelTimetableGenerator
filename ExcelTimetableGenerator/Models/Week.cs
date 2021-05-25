using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    [Table("ETG_Week")]
    public class Week
    {
        public int WeekID { get; set; }

        [StringLength(5)]
        public string AcademicYearID { get; set; }

        public int? WeekNum { get; set; }

        [StringLength(50)]
        public string WeekDesc { get; set; }

        [StringLength(255)]
        public string Notes { get; set; }

        [Display(Name = "Other Details")]
        [StringLength(255)]
        public string Notes2 { get; set; }
    }
}
