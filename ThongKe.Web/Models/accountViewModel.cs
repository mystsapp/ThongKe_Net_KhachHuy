using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class accountViewModel
    {
        public string username { get; set; }

        public string password { get; set; }

        public string hoten { get; set; }

        public string daily { get; set; }

        public string chinhanh { get; set; }

        public string role { get; set; }

        public bool doimatkhau { get; set; }

        public DateTime ngaydoimk { get; set; }

        public bool trangthai { get; set; }

        public string khoi { get; set; }

        public string nguoitao { get; set; }

        public DateTime ngaytao { get; set; }

        public string nguoicapnhat { get; set; }

        public DateTime? ngaycapnhat { get; set; }

        public string nhom { get; set; }
    }
}