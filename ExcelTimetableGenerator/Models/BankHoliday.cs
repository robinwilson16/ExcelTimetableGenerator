using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class BankHoliday
    {
        public int BankHolidayID { get; set; }

        [StringLength(255)]
        public string BankHolidayDesc { get; set; }
    }
}
