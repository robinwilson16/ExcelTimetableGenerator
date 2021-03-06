﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Course
    {
        public int CourseID { get; set; }

        public int ProgrammeID { get; set; }

        [StringLength(250)]
        public string SiteCode { get; set; }

        [StringLength(250)]
        public string SiteName { get; set; }

        [StringLength(72)]
        public string CourseCode { get; set; }

        [StringLength(255)]
        public string CourseTitle { get; set; }

        public bool? IsMainCourse { get; set; }

        [StringLength(50)]
        public string CourseStatus { get; set; }

        public int CourseOrder { get; set; }

        [StringLength(8)]
        public string AimCode { get; set; }

        [StringLength(150)]
        public string AwardBody { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime? StartDate { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime? EndDate { get; set; }

        public int? PLH1618 { get; set; }
        public int? PLH19 { get; set; }
        public int? PLHMax { get; set; }
        public int? EEP1618 { get; set; }
        public int? EEP19 { get; set; }
        public int? EEPMax { get; set; }

        [Display(Name = "Hours per Week")]
        [Column(TypeName = "decimal(9,2)")]
        public decimal? HoursPerWeek { get; set; }
        public int? Weeks { get; set; }
        public int? YearNo { get; set; }
        public int? NoOfYears { get; set; }

        [StringLength(250)]
        public string ModeOfAttendanceCode { get; set; }

        [StringLength(250)]
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
        public string Notes { get; set; }

        public Programme Programme { get; set; }
    }
}
