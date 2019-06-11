using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class StaffMember
    {
        [StringLength(50)]
        public string StaffRef { get; set; }

        [StringLength(50)]
        public string Surname { get; set; }

        [StringLength(50)]
        public string Forename { get; set; }

        [StringLength(150)]
        public string StaffDetails { get; set; }
    }
}
