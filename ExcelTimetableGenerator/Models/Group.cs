using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Group
    {
        [StringLength(20)]
        public string GroupID { get; set; }
        public int ProgrammeID { get; set; }

        [StringLength(2)]
        public string GroupCode { get; set; }

        [StringLength(75)]
        public string ProgCodeWithGroup { get; set; }

        public Programme Programme { get; set; }
    }
}
