using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    [Table("ETG_BankHoliday")]
    public class BankHoliday
    {
        public int BankHolidayID { get; set; }

        [StringLength(5)]
        public string AcademicYearID { get; set; }

        [StringLength(255)]
        public string BankHolidayDesc { get; set; }
    }
}
