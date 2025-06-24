using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CitipostWarehouseApplication.Models
{
    public class Subscriber
    {
        public string ContactFullName { get; set; }
        public string AccountName { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string City { get; set; }
        public string StateProvince { get; set; }
        public string PostCode { get; set; }
        public string Country { get; set; }
        public string AssignedSupplier { get; set; }
    }
}
