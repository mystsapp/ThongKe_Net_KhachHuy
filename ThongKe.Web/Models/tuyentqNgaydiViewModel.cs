using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class tuyentqNgaydiViewModel
    {
        public long stt { get; set; }
        public string chinhanh { get; set; }
        public string tuyentq { get; set; }
        public Nullable<int> khachht { get; set; }
        public Nullable<decimal> thucthuht { get; set; }
        public Nullable<int> khachcu { get; set; }
        public Nullable<decimal> thucthucu { get; set; }
    }
}