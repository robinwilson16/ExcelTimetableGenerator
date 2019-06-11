using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class SelectListData
    {
        [StringLength(10)]
        public string Code { get; set; }

        [StringLength(255)]
        public string Description { get; set; }
    }
}
