using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Models
{
    public class CustomerArea
    {
        public int CustomerID { get; set; }
        public int AreaID { get; set; }

        public Customer Customer { get; set; }
        public Area Area { get; set; }
    }
}
