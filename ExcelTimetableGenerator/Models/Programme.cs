using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Programme
    {
        public int ProgrammeID { get; set; }

        [StringLength(20)]
        public string SiteCode { get; set; }

        [StringLength(100)]
        public string SiteName { get; set; }

        [StringLength(24)]
        public string FacCode { get; set; }

        [StringLength(150)]
        public string FacName { get; set; }

        [StringLength(24)]
        public string TeamCode { get; set; }

        [StringLength(150)]
        public string TeamName { get; set; }

        [StringLength(72)]
        public string ProgCode { get; set; }

        [StringLength(255)]
        public string ProgTitle { get; set; }

        [StringLength(40)]
        public string ProgStatus { get; set; }

        public ICollection<Course> Course { get; set; }
    }
}
