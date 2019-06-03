using Microsoft.AspNetCore.Mvc.ModelBinding;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class Customer
    {
        [Display(Name = "Customer ID")]
        public int CustomerID { get; set; }

        [Display(Name = "Legacy ID")]
        public int? LegacyCustomerID { get; set; }

        [Display(Name = "Surname")]
        [StringLength(100)]
        [Required]
        public string Surname { get; set; }

        [Display(Name = "Forename")]
        [StringLength(100)]
        [Required]
        public string Forename { get; set; }

        [Display(Name = "Title")]
        [StringLength(10)]
        [BindRequired]
        [Required(ErrorMessage = "Select a title")]
        public string Title { get; set; }

        [Display(Name = "Email")]
        [StringLength(200)]
        [DataType(DataType.EmailAddress)]
        public string Email { get; set; }

        [Display(Name = "Mobile")]
        [DataType(DataType.PhoneNumber)]
        [StringLength(15)]
        public string MobilePhone { get; set; }

        [Display(Name = "Work Tel")]
        [DataType(DataType.PhoneNumber)]
        [StringLength(15)]
        public string WorkPhone { get; set; }

        [Display(Name = "Order Value")]
        [DataType(DataType.Currency)]
        [Column(TypeName = "decimal(18,2)")]
        public decimal? OrderValue { get; set; }

        [Display(Name = "Can Contact?")]
        public bool CanBeContacted { get; set; }

        [Display(Name = "Outstanding Remedial Work?")]
        public bool HasOutstandingRemedialWork { get; set; }

        [Display(Name = "Date of Enquiry")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime? DateOfEnquiry { get; set; }

        [Display(Name = "Showroom")]
        [BindRequired]
        [Required(ErrorMessage = "Select a showroom")]
        public int ShowroomID { get; set; }

        [Display(Name = "Source of Info")]
        public int? SourceOfInformationID { get; set; }

        [Display(Name = "Created Date")]
        public DateTime CreatedDate { get; set; }

        public string CreatedBy { get; set; }

        [Display(Name = "Updated Date")]
        public DateTime? UpdatedDate { get; set; }

        public string UpdatedBy { get; set; }

        public Showroom Showroom { get; set; }
        public ICollection<CustomerArea> CustomerArea { get; set; }
    }
}
