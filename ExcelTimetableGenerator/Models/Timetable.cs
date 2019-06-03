using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Timetable
    {
        public int TimetableID { get; set; }

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
        public string ProgName { get; set; }

        [StringLength(40)]
        public string ProgStatus { get; set; }

        [StringLength(72)]
        public string CourseCode { get; set; }

        [StringLength(255)]
        public string CourseTitle { get; set; }

        public bool? IsMainCourse { get; set; }

        [StringLength(40)]
        public string CourseStatus { get; set; }

        public int CourseOrder { get; set; }

        [StringLength(8)]
        public string AimCode { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime? StartDate { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime? EndDate { get; set; }

        public int? PLH1618 { get; set; }
        public int? PLH19 { get; set; }
        public int? EEP1618 { get; set; }
        public int? EEP19 { get; set; }

        [Display(Name = "Hours per Week")]
        [Column(TypeName = "decimal(9,2)")]
        public decimal? HoursPerWeek { get; set; }
        public int? Weeks { get; set; }
        public int? YearNo { get; set; }
        public int? NoOfYears { get; set; }

        [StringLength(2)]
        public string ModeOfAttendanceCode { get; set; }

        [StringLength(50)]
        public string ModeOfAttendanceName { get; set; }

        [StringLength(2)]
        public string FundStream { get; set; }

        [StringLength(3)]
        public string FundSource { get; set; }

        [StringLength(2)]
        public string ProgType { get; set; }

        [StringLength(20)]
        public string FundModel { get; set; }

        public int? GroupSize { get; set; }
        public int? NumGroups { get; set; }
    }
}
