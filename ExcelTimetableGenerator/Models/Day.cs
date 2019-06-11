using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Day
    {
        public int DayID { get; set; }

        [StringLength(20)]
        public string DayName { get; set; }

        public int Slots { get; set; }
    }
}
