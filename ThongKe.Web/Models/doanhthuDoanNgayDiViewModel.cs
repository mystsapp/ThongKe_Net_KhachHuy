using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class doanhthuDoanNgayDiViewModel
    {
        public long stt { get; set; }
        public string sgtcode { get; set; }
        public string tuyentq { get; set; }
        public Nullable<System.DateTime> batdau { get; set; }
        public Nullable<System.DateTime> ketthuc { get; set; }
        public Nullable<int> sokhach { get; set; }
        public Nullable<decimal> doanhthu { get; set; }
    }
}