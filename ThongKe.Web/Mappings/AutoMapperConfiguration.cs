using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ThongKe.Data.Models.EF;
using ThongKe.Web.Models;

namespace ThongKe.Web.Mappings
{
    public class AutoMapperConfiguration
    {
        public static void configure()
        {
            Mapper.CreateMap<account, accountViewModel>();
            Mapper.CreateMap<chinhanh, chinhanhViewModel>();
            Mapper.CreateMap<dmdaily, dmdailyViewModel>();

            Mapper.CreateMap<doanhthuSaleQuay, doanhthuSaleQuayViewModel>();
            Mapper.CreateMap<doanthuQuayNgayBan, doanthuQuayNgayBanViewModel>();

            Mapper.CreateMap<doanhthuDoanNgayDi, doanhthuDoanNgayDiViewModel>();
            Mapper.CreateMap<tuyentqNgaydi, tuyentqNgaydiViewModel>();
            Mapper.CreateMap<doanhthuToanhethong, doanhthuToanhethongViewModel>();

            Mapper.CreateMap<thongkeweb, thongkewebViewModel>();

        }
    }
}