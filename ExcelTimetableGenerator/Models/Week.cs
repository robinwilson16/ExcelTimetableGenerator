using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Week
    {
        public int WeekID { get; set; }

        public int? WeekNum { get; set; }

        [StringLength(50)]
        public string WeekDesc { get; set; }

        [StringLength(255)]
        public string Notes { get; set; }
    }
}
