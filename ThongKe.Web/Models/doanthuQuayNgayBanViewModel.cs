using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class doanthuQuayNgayBanViewModel
    {
        public long stt { get; set; }
        public string dailyxuatve { get; set; }
        public string chinhanh { get; set; }
        public Nullable<int> sokhach { get; set; }
        public Nullable<decimal> doanhso { get; set; }
        public Nullable<decimal> doanhthu { get; set; }
    }
}