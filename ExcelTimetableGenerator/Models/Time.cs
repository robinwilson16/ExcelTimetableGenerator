using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Time
    {
        public int TimeID { get; set; }

        public string TimeName { get; set; }

        [StringLength(2)]
        public string Hours { get; set; }

        [StringLength(2)]
        public string Mins { get; set; }
    }
}
