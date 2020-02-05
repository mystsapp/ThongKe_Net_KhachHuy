using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ThongKe.Common;
using ThongKe.Service;
using ThongKe.Web.Models;

namespace ThongKe.Web.Controllers
{
    public class LoginController : Controller
    {
        private IaccountService _accountService;

        public LoginController(IaccountService accountService)
        {
            _accountService = accountService;
        }
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult login(LoginViewModel model)
        {
            if (ModelState.IsValid)
            {
                var result = _accountService.login(model.username, model.password);
                if (result == 0)
                {
                    ModelState.AddModelError("", "Tài khoản này không tồn tại");
                }
                else if (result == 1)
                {
                    var userInfo = _accountService.GetById(model.username);
                    Session[CommonConstants.UserSession] = userInfo;

                    Session["username"] = userInfo.username;
                    Session["hoten"] = userInfo.hoten;
                    Session["chinhanh"] = userInfo.chinhanh;
                    Session["role"] = userInfo.role;
                    Session["daily"] = userInfo.daily;
                    Session["khoi"] = userInfo.khoi;
                    Session["nhom"] = userInfo.nhom;

                    bool doi = userInfo.doimatkhau;
                    if (doi)
                    {
                        return RedirectToAction("changepass/" + userInfo.username, "login");
                    }
                    DateTime ngaydoimk = userInfo.ngaydoimk;
                    int kq = (DateTime.Now.Month - ngaydoimk.Month) + 12 * (DateTime.Now.Year - ngaydoimk.Year);//?
                    if (kq >= 2)
                    {
                        return RedirectToAction("changepass/" + userInfo.username, "login");
                    }
                    return RedirectToAction("Index", "Home");
                }
                else if (result == -1)
                {
                    ModelState.AddModelError("", "Tài khoản này đã bị khóa");
                }
                else if (result == -2)
                {
                    ModelState.AddModelError("", "Mật khẩu không đúng.");
                }
            }
            return View("Index");
        }

        public ActionResult Logout()
        {
            Session.Clear();
            return Redirect("/");
        }

        [HttpGet]
        public ActionResult changepass(string name)
        {
            var entity = new changepassViewModel();
            var dao = _accountService.getUserByName(Session["username"].ToString());
            entity.username = dao.username;
            entity.password = dao.password;
            return View("changepass", entity);
        }
        [HttpPost]
        public ActionResult changepass(changepassViewModel model)
        {
            if (ModelState.IsValid)
            {
                var result = _accountService.changepass(model.username, model.password, model.newpassword, model.confirmpassword);
                if (result == -1)
                {
                    ModelState.AddModelError("", "Vui lòng nhập mật khẩu cũ.");
                }
                else if (result == -2)
                {
                    ModelState.AddModelError("", "Mật khẩu cũ không đúng.");
                }
                else if (result == -3)
                {
                    ModelState.AddModelError("", "Vui lòng nhập mật khẩu mới.");
                }
                else if (result == -4)
                {
                    ModelState.AddModelError("", "Vui lòng nhập lại mật khẩu mới.");
                }
                else if (result == -5)
                {
                    ModelState.AddModelError("", "Mật khẩu nhập lại không đúng.");
                }
                else if (result == 1)
                {
                    //if (Session["role"].Equals("cashier"))
                    //{
                    //    return RedirectToAction("Index", "Cashier");
                    //}
                    //else
                    //{
                    return RedirectToAction("Index", "Home");
                    //}
                }
            }
            return View("changepass");
        }

    }
}