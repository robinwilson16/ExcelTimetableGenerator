using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class TermDate
    {
        public int TermDateID { get; set; }

        [StringLength(50)]
        public string TermDateName { get; set; }

        public bool IsTerm { get; set; }

        [StringLength(255)]
        public string Dates { get; set; }
    }
}
