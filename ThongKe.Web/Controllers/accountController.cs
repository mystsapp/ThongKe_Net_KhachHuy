using AutoMapper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using ThongKe.Data.Models.EF;
using ThongKe.Service;
using ThongKe.Web.Infrastructure.Core;
using ThongKe.Web.Infrastructure.Extensions;
using ThongKe.Web.Models;

namespace ThongKe.Web.Controllers
{
    public class accountController : BaseController
    {
        private IaccountService _accountService;
        private ICommonService _commonService;

        public accountController(IaccountService accountService, ICommonService commonService)
        {
            _accountService = accountService;
            _commonService = commonService;
        }
        // GET: account
        public ActionResult Index()
        {

            return View();
        }

        [HttpGet]
        public JsonResult LoadData(string name, string status, int page, int pageSize)
        {
            int totalRow = 0;

            var listAccount = _accountService.Search(name, page, pageSize, status, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<account>, IEnumerable<accountViewModel>>(listAccount);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetDetail(string id)
        {
            bool statusUser = false;
            var user = _accountService.GetById(id);
            //var userViewModel = Mapper.Map<account, accountViewModel>(user);
            if (user != null)
            {
                statusUser = true;

            }
            return Json(new
            {
                data = user,
                status = statusUser
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult SaveData(string strUser, string Hidid, string hidPass)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            var user1 = serializer.Deserialize<accountViewModel>(strUser);

            var user = new account();
            user.Updateaccount(user1);

            bool status = true;
            string message = string.Empty;
            if (Hidid == "0")//dang them
            {
                user.password = _accountService.EncodeSHA1(user.password);
                user.ngaytao = DateTime.Now;
                user.ngaycapnhat = DateTime.Now;
                user.ngaydoimk=DateTime.Now;
                try
                {
                    _accountService.Add(user);
                    _accountService.Save();
                }
                catch (Exception ex)
                {
                    status = false;
                    //message = ex.Message;
                    throw ex;
                }
            }
            else if (Hidid != "0")
            {
                //var oldUser = _accountService.GetById(user.username);
                if (user.password != "") //password field is required
                {
                    user.password = _accountService.EncodeSHA1(user.password);
                    user.ngaydoimk = DateTime.Now;
                }
                else
                    user.password = hidPass;

                try
                {
                    _accountService.Update(user);
                    _accountService.Save();
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    status = false;
                    throw ex;
                }
            }
            return Json(new
            {
                status = status,
                message = message
            });
        }

        [HttpPost]
        public JsonResult Delete(string id)
        {
            bool status = true;
            string message = string.Empty;
            try
            {
                _accountService.Delete(id);
                _accountService.Save();
            }
            catch (Exception ex)
            {
                status = false;
                message = ex.Message;
            }

            return Json(new
            {
                status = status,
                message = message
            });
        }

        [HttpGet]
        public JsonResult GetAllChiNhanh()
        {
            var model = _commonService.GetAllChiNhanh();
            //var viewModel = Mapper.Map<IEnumerable<chinhanh>, IEnumerable<chinhanhViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(model)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetAllChiNhanhByNhom(string nhom,string chinhanh)
        {
            var model = _commonService.GetAllChiNhanhByNhom(nhom,chinhanh);
            //var viewModel = Mapper.Map<IEnumerable<chinhanh>, IEnumerable<chinhanhViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(model)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetAllTuyentqByKhoi(string khoi)
        {
            var model = _commonService.GetAllTuyentqByKhoi(khoi);
            //var viewModel = Mapper.Map<IEnumerable<chinhanh>, IEnumerable<chinhanhViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(model)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetDmdailyByChiNhanh(string chinhanh)
        {
            var model = _commonService.GetDmdailyByChiNhanh(chinhanh);
            var viewModel = Mapper.Map<IEnumerable<dmdaily>, IEnumerable<dmdailyViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(viewModel)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetDmdailyByNhomChiNhanh(string nhom)
        {
            var model = _commonService.GetDmdailyByNhomChiNhanh(nhom);
            var viewModel = Mapper.Map<IEnumerable<dmdaily>, IEnumerable<dmdailyViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(viewModel)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetAllDmDaiLy()
        {
            var model = _commonService.GetAllDmDaiLy();
            var viewModel = Mapper.Map<IEnumerable<dmdaily>, IEnumerable<dmdailyViewModel>>(model);
            return Json(new
            {
                data = JsonConvert.SerializeObject(viewModel)
            }, JsonRequestBehavior.AllowGet);
        }
    }
}