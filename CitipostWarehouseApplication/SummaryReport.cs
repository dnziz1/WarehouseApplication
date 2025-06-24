using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CitipostWarehouseApplication.Models
{
    public class SummaryReport
    {
        public string SupplierName { get; set; }
        public int TotalItems { get; set; }
        public List<string> Countries { get; set; } = new List<string>();
    }
}
