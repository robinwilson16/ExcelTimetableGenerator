using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    [Table("ETG_TermDate")]
    public class TermDate
    {
        public int TermDateID { get; set; }

        [StringLength(5)]
        public string AcademicYearID { get; set; }

        [StringLength(50)]
        public string TermDateName { get; set; }

        public bool IsTerm { get; set; }

        [StringLength(255)]
        public string Dates { get; set; }

        public int TermDateOrder { get; set; }
    }
}
