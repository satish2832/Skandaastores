using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StoreSite.Models
{
    public class CustomerEnquiry
    {
        public string CustomerName { get; set; }
        public string EmailAddress { get; set; }
        public string PhoneNumber { get; set; }
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string ProductCode { get; set; }
    }
}