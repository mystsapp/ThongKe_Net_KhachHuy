using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class doanhthuSaleQuayViewModel
    {
        public long stt { get; set; }
        public string nguoixuatve { get; set; }
        public Nullable<decimal> doanhso { get; set; }
        public Nullable<decimal> thucthu { get; set; }
        public string chinhanh { get; set; }
    }
}