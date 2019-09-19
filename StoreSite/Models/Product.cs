using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StoreSite.Models
{
    public class Product
    {
        public string Code { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public string Discount { get; set; }
        public string Variant { get; set; }
        public string Colors { get; set; }       
    }
}