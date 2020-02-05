using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ThongKe.Common.ViewModel;
using ThongKe.Data.Models.EF;
using ThongKe.Service;
using ThongKe.Web.Infrastructure.Core;
using ThongKe.Web.Models;

namespace ThongKe.Web.Controllers
{
    public class HomeController : BaseController
    {
        private IThongKeService _thongkeService;
        private ICommonService _commonService;

        public HomeController(IThongKeService thongKeService, ICommonService commonService)
        {
            _thongkeService = thongKeService;
            _commonService = commonService;
        }
        public ActionResult Index()
        {
            return View();
        }

        public JsonResult LoadDataThongKeSoKhachOB()
        {
            string khoi = "OB";

            var listOB = _commonService.ThongKeSoKhachOB(khoi);//doanhthuKhachLeHeThongEntities

            return Json(new
            {
                data = listOB,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadDataThongKeSoKhachND()
        {
            string khoi = "ND";

            var listOB = _commonService.ThongKeSoKhachOB(khoi);//doanhthuKhachLeHeThongEntities

            return Json(new
            {
                data = listOB,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadDataThongKeDoanhThuOB()
        {
            string khoi = "OB";

            var listOB = _commonService.ThongKeDoanhThuOB(khoi);//doanhthuKhachLeHeThongEntities

            return Json(new
            {
                data = listOB,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadDataThongKeDoanhThuND()
        {
            string khoi = "ND";

            var listOB = _commonService.ThongKeDoanhThuOB(khoi);//doanhthuKhachLeHeThongEntities

            return Json(new
            {
                data = listOB,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }
    }
}