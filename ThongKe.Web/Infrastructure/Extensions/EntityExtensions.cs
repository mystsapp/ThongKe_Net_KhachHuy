using ThongKe.Data.Models.EF;
using ThongKe.Web.Models;

namespace ThongKe.Web.Infrastructure.Extensions
{
    public static class EntityExtensions
    {
        public static void Updateaccount(this account acc, accountViewModel accViewModel)
        {
            acc.username = accViewModel.username;
            acc.password = accViewModel.password;
            acc.hoten = accViewModel.hoten;
            acc.daily = accViewModel.daily;
            acc.chinhanh = accViewModel.chinhanh;
            acc.role = accViewModel.role;
            acc.doimatkhau = accViewModel.doimatkhau;
            acc.ngaydoimk = accViewModel.ngaydoimk;
            acc.trangthai = accViewModel.trangthai;
            acc.khoi = accViewModel.khoi;
            acc.nguoitao = accViewModel.nguoitao;
            acc.ngaytao = accViewModel.ngaytao;

            acc.nguoicapnhat = accViewModel.nguoicapnhat;
            acc.ngaycapnhat = accViewModel.ngaycapnhat;

            acc.nhom = accViewModel.nhom;
        }
    }
}