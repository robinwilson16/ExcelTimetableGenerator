using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    [Table("ETG_DaySlot")]
    public class DaySlot
    {
        public int DaySlotID { get; set; }

        [StringLength(5)]
        public string AcademicYearID { get; set; }

        [StringLength(20)]
        public string DayName { get; set; }

        public int NumSlots { get; set; }
    }
}
