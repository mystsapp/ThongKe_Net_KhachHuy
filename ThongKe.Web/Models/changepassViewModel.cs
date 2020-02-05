using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ThongKe.Web.Models
{
    public class changepassViewModel
    {
        [Display(Name = "Tên đăng nhập")]
        public string username { get; set; }

        [Display(Name = "Mật khẩu")]
        [Required(ErrorMessage = "Vui lòng nhập password cũ")]
        public string password { get; set; }

        [Display(Name = "Mật khẩu mới")]
        [Required(ErrorMessage = "Vui lòng nhập password mới")]
        public string newpassword { get; set; }
        [Display(Name = "Mật khẩu mới")]

        [Required(ErrorMessage = "Vui lòng nhập lại password mới")]
        public string confirmpassword { get; set; }
    }
}