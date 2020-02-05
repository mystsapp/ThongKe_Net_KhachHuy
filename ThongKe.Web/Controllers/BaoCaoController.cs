using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using ThongKe.Data.Repositories;
using ThongKe.Service;
using ThongKe.Web.Infrastructure.Core;
using System.Web.Script.Serialization;
using ThongKe.Data.Models.EF;
using ThongKe.Web.Infrastructure.Extensions;
using ThongKe.Web.Models;
using Newtonsoft.Json;
using AutoMapper;
using System.Text;
using System.Text.RegularExpressions;

namespace ThongKe.Web.Controllers
{
    public class BaoCaoController : BaseController
    {
        private IThongKeService _thongkeService;
        private IaccountService _accountService;
        private ICommonService _commonService;
        public BaoCaoController(IThongKeService thongKeService)
        {
            _thongkeService = thongKeService;

        }
        // GET: BaoCao
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult SaleTheoQuay()
        {
            return View();
        }//

        public ActionResult KhachHuy()
        {
            return View();
        }//

        public ActionResult KinhDoanhOnline()
        {
            return View();
        }//
        public ActionResult KinhDoanhOnline_ngaydi()
        {
            return View();
        }
        [HttpPost]
        public ViewResult SaleTheoQuay(string tungay, string denngay, string chinhanh, string khoi)//(string tungay,string denngay, string daily)
        {
            chinhanh = String.IsNullOrEmpty(chinhanh) ? Session["chinhanh"].ToString() : chinhanh;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 50;// sales
            xlSheet.Column(3).Width = 10;//stt
            xlSheet.Column(4).Width = 30;// doanh so
            xlSheet.Column(5).Width = 30;// doanh thu sale

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY BÁN QUẦY " + khoi + " " + chinhanh;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 5].Merge = true;
            setCenterAligment(2, 1, 2, 5, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 5].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 5, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Nhân viên ";
            xlSheet.Cells[5, 3].Value = "Code chinhanh ";

            xlSheet.Cells[5, 4].Value = "Tổng tiền";
            xlSheet.Cells[5, 5].Value = "Doanh số";

            xlSheet.Cells[5, 1, 5, 5].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table  
            int dong = 5;


            DataTable dt = _thongkeService.doanhthuSaleTheoQuay(tungay, denngay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (6 + dt.Rows.Count - 1) + ")";


            setBorder(5, 1, 5 + dt.Rows.Count, 5, xlSheet);
            setFontBold(5, 1, 5, 5, 11, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 5, 11, xlSheet);
            // canh giua cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);
            // canh giua code chinhanh
            setCenterAligment(6, 3, 6 + dt.Rows.Count, 3, xlSheet);
            NumberFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            // định dạng số cot tong cong
            NumberFormat(dong, 4, dong, 5, xlSheet);
            setBorder(dong, 4, dong, 5, xlSheet);
            setFontBold(dong, 4, dong, 5, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay_" + Session["daily"].ToString() + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuSale_" + khoi + " " + chinhanh + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }

        public ViewResult SaleTheoNgayDi()
        {

            return View();
        }//

        [HttpPost]
        public ViewResult SaleTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi)
        {
            // cn = Session["chinhanh"].ToString();
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 50;// sales
            xlSheet.Column(3).Width = 10;// code cn
            xlSheet.Column(4).Width = 30;// doanh so
            xlSheet.Column(5).Width = 30;// doanh thu sale

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY ĐI SALE " + khoi + " " + chinhanh;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 5].Merge = true;
            setCenterAligment(2, 1, 2, 5, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 5].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 5, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Nhân viên ";
            xlSheet.Cells[5, 3].Value = "Code CN ";

            xlSheet.Cells[5, 4].Value = "Tổng tiền";
            xlSheet.Cells[5, 5].Value = "Doanh số";

            xlSheet.Cells[5, 1, 5, 5].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = new DataTable();
            dt = _thongkeService.doanhthuSaleTheoNgayDi(tungay, denngay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số
            NumberFormat(dong, 4, dong, 5, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 5, xlSheet);
            setFontBold(5, 1, 5, 5, 11, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 5, 11, xlSheet);
            // canh giua cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);
            // canh giua code cn
            setCenterAligment(6, 3, 6 + dt.Rows.Count, 3, xlSheet);
            NumberFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            // định dạng số cot tong cong
            NumberFormat(dong, 4, dong, 5, xlSheet);
            setBorder(dong, 4, dong, 5, xlSheet);
            setFontBold(dong, 4, dong, 5, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay_" + Session["daily"].ToString() + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuSale_" + khoi + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }


        public ViewResult QuayTheoNgayBan()
        {

            return View();
        }//

        [HttpPost]
        public ViewResult QuayTheoNgayBan(string tungay, string denngay, string cn, string khoi)
        {
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 40;// quay
            xlSheet.Column(3).Width = 10;// cn
            xlSheet.Column(4).Width = 10;// so khach
            xlSheet.Column(5).Width = 20;// doanh số
            xlSheet.Column(6).Width = 20;// doanh thu

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY BÁN QUẦY " + khoi + "  " + cn;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 6].Merge = true;
            setCenterAligment(2, 1, 2, 6, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 6].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 6, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Văn phòng xuất vé ";
            xlSheet.Cells[5, 3].Value = "Code CN ";
            xlSheet.Cells[5, 4].Value = "Số khách";
            xlSheet.Cells[5, 5].Value = "Tổng tiền";
            xlSheet.Cells[5, 6].Value = "Doanh số";
            xlSheet.Cells[5, 1, 5, 6].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.doanhthuQuayTheoNgayBan(tungay, denngay, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {

                        //if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        //{
                        //    xlSheet.Cells[dong, j + 1].Value = "";
                        //}
                        //else
                        //{
                        xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        //}
                    }
                }

            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;

            // Sum tổng tiền
            xlSheet.Cells[dong, 3].Value = "TC";
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            // định dạng số
            NumberFormat(dong, 5, dong, 6, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 6, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 6, xlSheet);
            // font bold tieu de bang
            setFontBold(5, 1, 5, 6, 12, xlSheet);
            // font bold dong cuoi cùng
            setFontBold(dong, 1, dong, 6, 12, xlSheet);
            setBorder(dong, 3, dong, 6, xlSheet);
            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            // canh giưa cot chinhanh va so khach
            setCenterAligment(6, 3, 6 + dt.Rows.Count, 4, xlSheet);
            // dinh dạng number cot sokhach, doanh so, thuc thu
            NumberFormat(6, 5, 6 + dt.Rows.Count, 6, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay" + khoi + " " + cn + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }



        public ActionResult QuayTheoNgayDi()
        {
            return View();
        }//


        [HttpPost]
        public ViewResult QuayTheoNgayDi(string tungay, string denngay, string cn, string khoi)
        {
            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 40;// quay
            xlSheet.Column(3).Width = 10;// cn
            xlSheet.Column(4).Width = 10;// so khach
            xlSheet.Column(5).Width = 20;// doanh số
            xlSheet.Column(6).Width = 20;// doanh thu

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY ĐI QUẦY " + khoi + " " + cn;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 6].Merge = true;
            setCenterAligment(2, 1, 2, 6, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 6].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 6, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Văn phòng xuất vé ";
            xlSheet.Cells[5, 3].Value = "Code CN ";
            xlSheet.Cells[5, 4].Value = "Số khách";
            xlSheet.Cells[5, 5].Value = "Tổng tiền";
            xlSheet.Cells[5, 6].Value = "Doanh số";
            xlSheet.Cells[5, 1, 5, 6].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.doanhthuQuayTheoNgayDi(tungay, denngay, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        //if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        //{
                        //    xlSheet.Cells[dong, j + 1].Value = "";
                        //}
                        //else
                        //{
                        xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        //}
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 3].Value = "TC";
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            // định dạng số
            NumberFormat(dong, 5, dong, 6, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 6, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 6, xlSheet);
            setFontBold(5, 1, 5, 6, 12, xlSheet);
            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 3, dong, 6, xlSheet);
            setFontBold(dong, 1, dong, 6, 12, xlSheet);
            // canh giưa cot chinhanh va so khach
            setCenterAligment(6, 3, 6 + dt.Rows.Count, 4, xlSheet);
            // dinh dạng number cot sokhach, doanh so, thuc thu
            NumberFormat(6, 5, 6 + dt.Rows.Count, 6, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }

        public ActionResult DoanTheoNgayDi()
        {
            return View();
        }//

        [HttpPost]
        public ViewResult DoanTheoNgayDi(string tungay, string denngay, string cn, string khoi)
        {
            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// sgtcode
            xlSheet.Column(3).Width = 40;// tuyen tq
            xlSheet.Column(4).Width = 20;// bat dau 
            xlSheet.Column(5).Width = 20;// ket thu
            xlSheet.Column(6).Width = 10;// so khach
            xlSheet.Column(7).Width = 25;//doanh thu

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO ĐOÀN  " + khoi + "  " + cn;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 7].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Sgt Code ";
            xlSheet.Cells[5, 3].Value = "Tuyến tham quan ";
            xlSheet.Cells[5, 4].Value = "Ngày đi";
            xlSheet.Cells[5, 5].Value = "Ngày về";
            xlSheet.Cells[5, 6].Value = "Số khách";
            xlSheet.Cells[5, 7].Value = "Doanh số bán";
            xlSheet.Cells[5, 1, 5, 7].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _thongkeService.doanhthuDoanTheoNgay(tungay, denngay, cn, khoi);

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TC";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số
            NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 7, xlSheet);
            setFontBold(5, 1, 5, 6, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 7, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 7, xlSheet);
            setFontBold(dong, 5, dong, 7, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 7, 6 + dt.Rows.Count, 7, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuDoan_" + khoi + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }

        public ActionResult TuyentqTheoNgayDi()
        {

            return View();
        }

        [HttpPost]
        public ViewResult TuyentqTheoNgayDi(string tungay, string denngay, string cn, string khoi)
        {
            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 10;// chi nhanh
            xlSheet.Column(3).Width = 40;// tuyen tq
            xlSheet.Column(4).Width = 10;// sk ht
            xlSheet.Column(5).Width = 20;// doanh thu ht
            xlSheet.Column(6).Width = 10;// sk nam truoc
            xlSheet.Column(7).Width = 20;//doanh thu nam truoc
            xlSheet.Column(8).Width = 15;// ti le khach
            xlSheet.Column(9).Width = 15;// so sanh doanh thu

            xlSheet.Cells[2, 1].Value = "TUYẾN THAM QUAN THEO NGÀY ĐI TOUR " + Session["chinhanh"].ToString();
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 9].Merge = true;
            setCenterAligment(2, 1, 2, 9, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 9].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 9, xlSheet);

            // Tạo header
            // Tạo header

            xlSheet.Cells[5, 1].Value = "STT ";
            xlSheet.Cells[5, 1, 6, 1].Merge = true;
            xlSheet.Cells[5, 2].Value = "CN";
            xlSheet.Cells[5, 2, 6, 2].Merge = true;
            xlSheet.Cells[5, 3].Value = "Tuyến tham quan ";
            xlSheet.Cells[5, 3, 6, 3].Merge = true;

            xlSheet.Cells[5, 4].Value = "Thời điểm thống kê";
            xlSheet.Cells[5, 4, 5, 5].Merge = true;


            xlSheet.Cells[5, 6].Value = "So sánh cùng kỳ";
            xlSheet.Cells[5, 6, 5, 7].Merge = true;

            xlSheet.Cells[5, 8].Value = "Tỉ lệ % tăng giảm ";
            xlSheet.Cells[5, 8, 5, 9].Merge = true;
            // dong thu 2
            xlSheet.Cells[6, 4].Value = "Số khách";
            xlSheet.Cells[6, 5].Value = "Doanh số";
            xlSheet.Cells[6, 6].Value = "Số khách";
            xlSheet.Cells[6, 7].Value = "Doanh số";
            xlSheet.Cells[6, 8].Value = "Số khách";
            xlSheet.Cells[6, 9].Value = "Doanh số";
            setCenterAligment(5, 1, 6, 9, xlSheet);
            xlSheet.Cells[5, 1, 6, 9].Style.Font.SetFromFont(new Font("Times New Roman", 11, FontStyle.Bold));


            xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 11, FontStyle.Bold));

            // do du lieu tu table
            int dong = 6;


            DataTable dt = _thongkeService.doanhthuTuyentqTheoNgay(tungay, denngay, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                        if (Convert.ToInt32(xlSheet.Cells[dong, 4].Value) == 0 || Convert.ToInt32(xlSheet.Cells[dong, 6].Value) == 0)
                        {
                            xlSheet.Cells[dong, 8].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
                        }
                        if (Convert.ToDouble(xlSheet.Cells[dong, 5].Value) == 0 || Convert.ToDouble(xlSheet.Cells[dong, 7].Value) == 0)
                        {
                            xlSheet.Cells[dong, 9].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            // phan tram tong
            xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
            xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (7 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (7 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (7 + dt.Rows.Count - 1) + ")";
            PhantramFormat(6, 8, 6 + dt.Rows.Count + 1, 9, xlSheet);
            // định dạng số
            NumberFormat(dong, 4, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count + 2, 9, xlSheet);
            setFontBold(5, 1, 5, 5, 12, xlSheet);
            setFontSize(7, 1, 6 + dt.Rows.Count + 2, 9, 12, xlSheet);
            // dinh dang giu cho so khach
            setCenterAligment(7, 1, 7 + dt.Rows.Count, 2, xlSheet);
            setCenterAligment(7, 4, 7 + dt.Rows.Count, 4, xlSheet);
            setCenterAligment(7, 6, 7 + dt.Rows.Count, 6, xlSheet);
            setCenterAligment(7, 8, 7 + dt.Rows.Count, 9, xlSheet);
            // dinh dạng number cot sokhach, doanh so, thuc thu
            NumberFormat(7, 5, 7 + dt.Rows.Count + 1, 5, xlSheet);
            NumberFormat(7, 7, 6 + dt.Rows.Count + 1, 7, xlSheet);


            setBorder(dong, 4, dong, 9, xlSheet);
            setFontBold(dong, 4, dong, 9, 12, xlSheet);

            //xlSheet.View.FreezePanes(7, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuTuyentq" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();

        }

        public ActionResult TuyentqTheoQuy()
        {

            return View();
        }

        [HttpPost]
        public ViewResult TuyentqTheoQuy(int quy, int nam, string cn, string khoi)
        {
            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            int thang = 1;
            switch (quy)
            {
                case 1:
                    thang = 1;
                    break;
                case 2:
                    thang = 4;
                    break;
                case 3:
                    thang = 7;
                    break;
                case 4:
                    thang = 10;
                    break;
                default:
                    thang = 1;
                    break;

            }

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 50;// tuyen tq
            xlSheet.Column(3).Width = 10;// sk thang 1 nam hien tai
            xlSheet.Column(4).Width = 15;// doanh so thang 1 nam hien tai
            xlSheet.Column(5).Width = 10;// sk thang 1 nam trươc
            xlSheet.Column(6).Width = 15;// doanh so thang 1 nam truoc

            xlSheet.Column(7).Width = 10;// sk thang 2 nam hien tai
            xlSheet.Column(8).Width = 15;// doanh so thang 2 nam hien tai
            xlSheet.Column(9).Width = 10;// sk thang 2 nam trươc
            xlSheet.Column(10).Width = 15;// doanh so thang 2 nam truoc

            xlSheet.Column(11).Width = 10;// sk thang 3 nam hien tai
            xlSheet.Column(12).Width = 15;// doanh so thang 3 nam hien tai
            xlSheet.Column(13).Width = 10;// sk thang 3 nam trươc
            xlSheet.Column(14).Width = 15;// doanh so thang 3 nam truoc

            xlSheet.Cells[2, 1].Value = "THỐNG KÊ TUYẾN TQ THEO QUÝ " + quy + " NĂM " + nam+ " "+cn;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 14].Merge = true;
            setCenterAligment(2, 1, 2, 14, xlSheet);
            // dinh dang tu ngay den ngay

            //xlSheet.Cells[4, 1].Value = "";
            //xlSheet.Cells[3, 1, 3, 6].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 6, xlSheet);

            // Tạo header
            xlSheet.Cells[4, 1, 6, 14].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            xlSheet.Cells[4, 1].Value = "STT";
            xlSheet.Cells[4, 1, 6, 1].Merge = true;
            xlSheet.Cells[4, 2].Value = "Tuyến tham quan ";
            xlSheet.Cells[4, 2, 6, 2].Merge = true;
            // thang thứ nhất
            xlSheet.Cells[4, 3].Value = "Tháng " + thang;
            xlSheet.Cells[4, 3, 4, 6].Merge = true;
            // nam hiên tại của tháng thứ nhất
            xlSheet.Cells[5, 3].Value = nam;
            xlSheet.Cells[5, 3, 5, 4].Merge = true;
            // năm trước đó của tháng thứ nhất
            xlSheet.Cells[5, 5].Value = (nam - 1).ToString();
            xlSheet.Cells[5, 5, 5, 6].Merge = true;
            xlSheet.Cells[6, 3].Value = "SK";
            xlSheet.Cells[6, 4].Value = "Doanh số";
            // so khach va doanh so năm trước tháng 1
            xlSheet.Cells[6, 5].Value = "SK";
            xlSheet.Cells[6, 6].Value = "Doanh số";

            // thang thứ hai
            xlSheet.Cells[4, 7].Value = "Tháng " + (thang + 1).ToString();
            xlSheet.Cells[4, 7, 4, 10].Merge = true;
            // nam hiên tại của tháng thứ hai
            xlSheet.Cells[5, 7].Value = nam;
            xlSheet.Cells[5, 7, 5, 8].Merge = true;
            // năm trước đó của tháng thứ hai
            xlSheet.Cells[5, 9].Value = (nam - 1).ToString();
            xlSheet.Cells[5, 9, 5, 10].Merge = true;
            xlSheet.Cells[6, 7].Value = "SK";
            xlSheet.Cells[6, 8].Value = "Doanh số";
            // so khach va doanh so năm trước tháng 1
            xlSheet.Cells[6, 9].Value = "SK";
            xlSheet.Cells[6, 10].Value = "Doanh số";


            // thang thứ ba
            xlSheet.Cells[4, 11].Value = "Tháng " + (thang + 2).ToString();
            xlSheet.Cells[4, 11, 4, 14].Merge = true;
            // nam hiên tại của tháng thứ ba
            xlSheet.Cells[5, 11].Value = nam;
            xlSheet.Cells[5, 11, 5, 12].Merge = true;
            // năm trước đó của tháng thứ nhất
            xlSheet.Cells[5, 13].Value = (nam - 1).ToString();
            xlSheet.Cells[5, 13, 5, 14].Merge = true;
            xlSheet.Cells[6, 11].Value = "SK";
            xlSheet.Cells[6, 12].Value = "Doanh số";
            // so khach va doanh so năm trước tháng 1
            xlSheet.Cells[6, 13].Value = "SK";
            xlSheet.Cells[6, 14].Value = "Doanh số";
            // canh giữa cho tiêu đề bảng
            setCenterAligment(4, 1, 6, 14, xlSheet);

            // do du lieu tu table
            int dong = 6;


            DataTable dt = new DataTable();
            dt = _thongkeService.doanhthuTuyentqTheoQuy(quy, nam, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;


            setBorder(4, 1, 4 + dt.Rows.Count + 2, 14, xlSheet);
            setFontSize(7, 1, 7 + dt.Rows.Count, 14, 11, xlSheet);

            // định dạng number cho cột doanh số
            NumberFormat(7, 3, 7 + dt.Rows.Count+1, 14, xlSheet);

            // canh giua cot stt
            setCenterAligment(7, 1, 7 + dt.Rows.Count, 1, xlSheet);
            // canh giua so khach thang 1
            setCenterAligment(7, 3, 7 + dt.Rows.Count, 3, xlSheet);
            setCenterAligment(7, 5, 7 + dt.Rows.Count, 5, xlSheet);
            setCenterAligment(7, 7, 7 + dt.Rows.Count, 7, xlSheet);
            setCenterAligment(7, 9, 7 + dt.Rows.Count, 9, xlSheet);
            setCenterAligment(7, 11, 7 + dt.Rows.Count, 11, xlSheet);
            setCenterAligment(7, 13, 7 + dt.Rows.Count, 13, xlSheet);
            //
            xlSheet.Cells[dong, 3].Formula = "SUM(C7:C" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 4].Formula = "SUM(D7:D" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E7:E" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F7:F" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 7].Formula = "SUM(G7:G" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 8].Formula = "SUM(H7:H" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 9].Formula = "SUM(I7:I" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 10].Formula = "SUM(J7:J" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 11].Formula = "SUM(K7:K" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 12].Formula = "SUM(L7:L" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 13].Formula = "SUM(M7:M" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 14].Formula = "SUM(N7:N" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 3, dong, 14].Style.Font.SetFromFont(new Font("Times New Roman", 11, FontStyle.Bold));
            setBorder(dong, 3, dong,14, xlSheet);
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay_" + Session["daily"].ToString() + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "ThongkeTuyentqTheoQuy" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();

        }


        public ActionResult SaleTheoTuyentq()
        {

            return View();
        }//

        [HttpPost]
        public ViewResult SaleTheoTuyentq(string tungay, string denngay, string tuyentq, string khoi)
        {
            // cn = Session["chinhanh"].ToString();
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 50;// sales
            xlSheet.Column(3).Width = 10;// code cn
            xlSheet.Column(4).Width = 50;// tuyentq
            xlSheet.Column(5).Width = 20;// doanh so
            xlSheet.Column(6).Width = 20;// doanh thu sale

            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU SALE THEO TUYẾN " + tuyentq;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 6].Merge = true;
            setCenterAligment(2, 1, 2, 6, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 6].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 6, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Nhân viên ";
            xlSheet.Cells[5, 3].Value = "Code CN ";
            xlSheet.Cells[5, 4].Value = "Tuyến tham quan";
            xlSheet.Cells[5, 5].Value = "Tổng tiền";
            xlSheet.Cells[5, 6].Value = "Doanh số";

            xlSheet.Cells[5, 1, 5, 5].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            
            DataTable dt = new DataTable();
            dt = _thongkeService.doanhthuSaleTheoTuyettq(tungay, denngay, tuyentq, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số
            NumberFormat(dong, 5, dong, 6, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 6, xlSheet);
            setFontBold(5, 1, 5, 6, 11, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 5, 11, xlSheet);
            // canh giua cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);
            // canh giua code cn
            setCenterAligment(6, 3, 6 + dt.Rows.Count, 3, xlSheet);
            NumberFormat(6, 5, 6 + dt.Rows.Count, 6, xlSheet);
            // định dạng số cot tong cong
            //NumberFormat(dong, 4, dong, 5, xlSheet);
            setBorder(dong, 5, dong, 6, xlSheet);
            setFontBold(dong, 5, dong, 6, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay_" + Session["daily"].ToString() + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuSaleTheoTuyen_" + khoi + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }


        //public ActionResult TheoTuyentq()
        //{

        //    return View();
        //}
        //[HttpPost]
        //public ViewResult TheoTuyentq(string tungay, string denngay, string cn, string khoi)
        //{
        //    cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
        //    khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
        //    string fromTo = "";
        //    ExcelPackage ExcelApp = new ExcelPackage();
        //    ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
        //    // Định dạng chiều dài cho cột
        //    xlSheet.Column(1).Width = 10;//STT
        //    xlSheet.Column(2).Width = 10;// chi nhanh
        //    xlSheet.Column(3).Width = 40;// tuyen tq
        //    xlSheet.Column(4).Width = 10;// sk ht
        //    xlSheet.Column(5).Width = 20;// doanh thu ht
        //    xlSheet.Column(6).Width = 10;// sk nam truoc
        //    xlSheet.Column(7).Width = 20;//doanh thu nam truoc
        //    xlSheet.Column(8).Width = 15;// ti le khach
        //    xlSheet.Column(9).Width = 15;// so sanh doanh thu

        //    xlSheet.Cells[2, 1].Value = "TUYẾN THAM QUAN THEO NGÀY ĐI TOUR " + Session["chinhanh"].ToString();
        //    xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
        //    xlSheet.Cells[2, 1, 2, 9].Merge = true;
        //    setCenterAligment(2, 1, 2, 9, xlSheet);
        //    // dinh dang tu ngay den ngay
        //    if (tungay == denngay)
        //    {
        //        fromTo = "Ngày: " + tungay;
        //    }
        //    else
        //    {
        //        fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
        //    }
        //    xlSheet.Cells[3, 1].Value = fromTo;
        //    xlSheet.Cells[3, 1, 3, 9].Merge = true;
        //    xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
        //    setCenterAligment(3, 1, 3, 9, xlSheet);

        //    // Tạo header
        //    // Tạo header

        //    xlSheet.Cells[5, 1].Value = "STT ";
        //    xlSheet.Cells[5, 1, 6, 1].Merge = true;
        //    xlSheet.Cells[5, 2].Value = "CN";
        //    xlSheet.Cells[5, 2, 6, 2].Merge = true;
        //    xlSheet.Cells[5, 3].Value = "Tuyến tham quan ";
        //    xlSheet.Cells[5, 3, 6, 3].Merge = true;

        //    xlSheet.Cells[5, 4].Value = "Thời điểm thống kê";
        //    xlSheet.Cells[5, 4, 5, 5].Merge = true;


        //    xlSheet.Cells[5, 6].Value = "So sánh cùng kỳ";
        //    xlSheet.Cells[5, 6, 5, 7].Merge = true;

        //    xlSheet.Cells[5, 8].Value = "Tỉ lệ % tăng giảm ";
        //    xlSheet.Cells[5, 8, 5, 9].Merge = true;
        //    // dong thu 2
        //    xlSheet.Cells[6, 4].Value = "Số khách";
        //    xlSheet.Cells[6, 5].Value = "Doanh số";
        //    xlSheet.Cells[6, 6].Value = "Số khách";
        //    xlSheet.Cells[6, 7].Value = "Doanh số";
        //    xlSheet.Cells[6, 8].Value = "Số khách";
        //    xlSheet.Cells[6, 9].Value = "Doanh số";
        //    setCenterAligment(5, 1, 6, 9, xlSheet);
        //    xlSheet.Cells[5, 1, 6, 9].Style.Font.SetFromFont(new Font("Times New Roman", 11, FontStyle.Bold));


        //    xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 11, FontStyle.Bold));

        //    // do du lieu tu table
        //    int dong = 6;


        //    DataTable dt = _thongkeService.doanhthuTuyentqTheoNgay(tungay, denngay, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

        //    if (dt != null)
        //    {
        //        for (int i = 0; i < dt.Rows.Count; i++)
        //        {
        //            dong++;
        //            for (int j = 0; j < dt.Columns.Count; j++)
        //            {
        //                if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
        //                {
        //                    xlSheet.Cells[dong, j + 1].Value = 0;
        //                }
        //                else
        //                {
        //                    xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
        //                }
        //                if (Convert.ToInt32(xlSheet.Cells[dong, 4].Value) == 0 || Convert.ToInt32(xlSheet.Cells[dong, 6].Value) == 0)
        //                {
        //                    xlSheet.Cells[dong, 8].Value = 0;
        //                }
        //                else
        //                {
        //                    xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
        //                }
        //                if (Convert.ToDouble(xlSheet.Cells[dong, 5].Value) == 0 || Convert.ToDouble(xlSheet.Cells[dong, 7].Value) == 0)
        //                {
        //                    xlSheet.Cells[dong, 9].Value = 0;
        //                }
        //                else
        //                {
        //                    xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
        //                }
        //            }
        //        }
        //    }
        //    else
        //    {
        //        SetAlert("No sale.", "warning");
        //        return View();
        //    }

        //    dong++;
        //    // phan tram tong
        //    xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
        //    xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
        //    xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (7 + dt.Rows.Count - 1) + ")";
        //    xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (7 + dt.Rows.Count - 1) + ")";
        //    xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (7 + dt.Rows.Count - 1) + ")";
        //    xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (7 + dt.Rows.Count - 1) + ")";
        //    //xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (7 + dt.Rows.Count - 1) + ")";
        //    //xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (7 + dt.Rows.Count - 1) + ")";
        //    PhantramFormat(6, 8, 6 + dt.Rows.Count + 1, 9, xlSheet);
        //    // định dạng số
        //    NumberFormat(dong, 4, dong, 7, xlSheet);

        //    setBorder(5, 1, 5 + dt.Rows.Count + 2, 9, xlSheet);
        //    setFontBold(5, 1, 5, 5, 12, xlSheet);
        //    setFontSize(7, 1, 6 + dt.Rows.Count + 2, 9, 12, xlSheet);
        //    // dinh dang giu cho so khach
        //    setCenterAligment(7, 1, 7 + dt.Rows.Count, 2, xlSheet);
        //    setCenterAligment(7, 4, 7 + dt.Rows.Count, 4, xlSheet);
        //    setCenterAligment(7, 6, 7 + dt.Rows.Count, 6, xlSheet);
        //    setCenterAligment(7, 8, 7 + dt.Rows.Count, 9, xlSheet);
        //    // dinh dạng number cot sokhach, doanh so, thuc thu
        //    NumberFormat(7, 5, 7 + dt.Rows.Count + 1, 5, xlSheet);
        //    NumberFormat(7, 7, 6 + dt.Rows.Count + 1, 7, xlSheet);


        //    setBorder(dong, 4, dong, 9, xlSheet);
        //    setFontBold(dong, 4, dong, 9, 12, xlSheet);

        //    //xlSheet.View.FreezePanes(7, 20);


        //    Response.Clear();
        //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuTuyentq" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
        //    Response.BinaryWrite(ExcelApp.GetAsByteArray());
        //    Response.End();

        //    return View();

        //}
        public ViewResult KhachleHethong()
        {

            return View();
        }//

        [HttpPost]
        public ViewResult KhachLeHethong(string tungay, string denngay, string cn, string khoi)
        {

            cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("lienketkhachle");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 10;// cn
            xlSheet.Column(3).Width = 40;// quay
            xlSheet.Column(4).Width = 10;// so khach hien tai
            xlSheet.Column(5).Width = 20;// doanh số hien tai
            xlSheet.Column(6).Width = 10;// so khach nam truoc
            xlSheet.Column(7).Width = 20; // doanh thu nam truoc
            xlSheet.Column(8).Width = 15; // ti le so khach
            xlSheet.Column(9).Width = 15;// doanh thu so sanh

            xlSheet.Cells[2, 1].Value = "LIÊN KẾT QUẦY KHÁCH LẼ HỆ THỐNG " + khoi + "  " + cn;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 9].Merge = true;
            setCenterAligment(2, 1, 2, 9, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 9].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 9, xlSheet);

            // Tạo header

            xlSheet.Cells[5, 1].Value = "STT ";
            xlSheet.Cells[5, 1, 6, 1].Merge = true;
            xlSheet.Cells[5, 2].Value = "CN";
            xlSheet.Cells[5, 2, 6, 2].Merge = true;
            xlSheet.Cells[5, 3].Value = "Văn phòng xuất vé ";
            xlSheet.Cells[5, 3, 6, 3].Merge = true;

            xlSheet.Cells[5, 4].Value = "Thời điểm thống kê";
            xlSheet.Cells[5, 4, 5, 5].Merge = true;


            xlSheet.Cells[5, 6].Value = "So sánh cùng kỳ";
            xlSheet.Cells[5, 6, 5, 7].Merge = true;

            xlSheet.Cells[5, 8].Value = "Tỉ lệ % tăng giảm ";
            xlSheet.Cells[5, 8, 5, 9].Merge = true;
            // dong thu 2
            xlSheet.Cells[6, 4].Value = "Số khách";
            xlSheet.Cells[6, 5].Value = "Doanh số";
            xlSheet.Cells[6, 6].Value = "Số khách";
            xlSheet.Cells[6, 7].Value = "Doanh số";
            xlSheet.Cells[6, 8].Value = "Số khách";
            xlSheet.Cells[6, 9].Value = "Doanh số";
            setCenterAligment(5, 1, 6, 9, xlSheet);
            xlSheet.Cells[5, 1, 6, 9].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 6;

            DataTable dt = _thongkeService.doanhthuKhachleHethong(tungay, denngay, cn, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                        if (Convert.ToInt32(xlSheet.Cells[dong, 4].Value) == 0 || Convert.ToInt32(xlSheet.Cells[dong, 6].Value) == 0)
                        {
                            xlSheet.Cells[dong, 8].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
                        }
                        if (Convert.ToDouble(xlSheet.Cells[dong, 5].Value) == 0 || Convert.ToDouble(xlSheet.Cells[dong, 7].Value) == 0)
                        {
                            xlSheet.Cells[dong, 9].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
                        }

                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            // phan tram tong
            xlSheet.Cells[dong, 8].Formula = "=(" + (xlSheet.Cells[dong, 4]).Address + "-" + (xlSheet.Cells[dong, 6]).Address + ")/" + (xlSheet.Cells[dong, 6]).Address;
            xlSheet.Cells[dong, 9].Formula = "=(" + (xlSheet.Cells[dong, 5]).Address + "-" + (xlSheet.Cells[dong, 7]).Address + ")/" + (xlSheet.Cells[dong, 7]).Address;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 4].Formula = "SUM(D6:D" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 5].Formula = "SUM(E6:E" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (7 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (7 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (7 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (7 + dt.Rows.Count - 1) + ")";
            PhantramFormat(6, 8, 6 + dt.Rows.Count + 1, 9, xlSheet);
            // định dạng số
            NumberFormat(dong, 4, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count + 2, 9, xlSheet);
            setFontBold(5, 1, 5, 5, 12, xlSheet);
            setFontSize(7, 1, 6 + dt.Rows.Count + 2, 9, 12, xlSheet);
            // dinh dang giu cho so khach
            setCenterAligment(7, 1, 7 + dt.Rows.Count, 2, xlSheet);
            setCenterAligment(7, 4, 7 + dt.Rows.Count, 4, xlSheet);
            setCenterAligment(7, 6, 7 + dt.Rows.Count, 6, xlSheet);
            setCenterAligment(7, 8, 7 + dt.Rows.Count, 9, xlSheet);

            // dinh dạng number cot sokhach, doanh so, thuc thu
            NumberFormat(7, 5, 7 + dt.Rows.Count + 1, 5, xlSheet);
            NumberFormat(7, 7, 6 + dt.Rows.Count + 1, 7, xlSheet);


            setBorder(dong, 4, dong, 9, xlSheet);
            setFontBold(dong, 4, dong, 9, 12, xlSheet);

            //xlSheet.View.FreezePanes(7, 20);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "LienketKhachle" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();

        }


        //////////////////////////////////////////////////////////////////////////////////////////////
        [HttpGet]
        public JsonResult LoadDataSaleTheoQuay(string tungay, string denngay, string cn, string khoi, int page, int pageSize)
        {
            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            if (cn == null)
            {
                cn = "";
            }

            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            int totalRow = 0;

            var listAccount = _thongkeService.doanhthuSaleTheoQuayEntities(tungay, denngay, cn, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<doanhthuSaleQuay>, IEnumerable<doanhthuSaleQuayViewModel>>(listAccount);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult LoadDataSaleTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuSaleTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<doanhthuSaleQuay>, IEnumerable<doanhthuSaleQuayViewModel>>(listDoanhthu);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult LoadDataSaleTheoTuyentq(string tungay, string denngay, string tuyentq, string khoi, int page, int pageSize)
        {
            int totalRow = 0;
            tuyentq = tuyentq.Trim();
            var listDoanhthu = _thongkeService.doanhthuSaleTheoTuyentqEntities(tungay, denngay, tuyentq, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            //var responseData = Mapper.Map<IEnumerable<doanhthuSaleQuay>, IEnumerable<doanhthuSaleQuayViewModel>>(listDoanhthu);

            return Json(new
            {
                data = listDoanhthu,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }


        public JsonResult LoadDataQuayTheoNgayBan(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuQuayTheoNgayBanEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
           // var responseData = Mapper.Map<IEnumerable<doanthuQuayNgayBan>, IEnumerable<doanthuQuayNgayBanViewModel>>(listDoanhthu);

            return Json(new
            {
                data = listDoanhthu,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult LoadDataQuayTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuQuayTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<doanthuQuayNgayBan>, IEnumerable<doanthuQuayNgayBanViewModel>>(listDoanhthu);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult LoadDataDoanTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuDoanTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<doanhthuDoanNgayDi>, IEnumerable<doanhthuDoanNgayDiViewModel>>(listDoanhthu);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult LoadDataTuyentqTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuTuyentqTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<tuyentqNgaydi>, IEnumerable<tuyentqNgaydiViewModel>>(listDoanhthu);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadDataKhachLeHethong(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize)
        {
            int totalRow = 0;

            var listDoanhthu = _thongkeService.doanhthuKhachLeHeThongEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<doanhthuToanhethong>, IEnumerable<doanhthuToanhethongViewModel>>(listDoanhthu);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        //////////////////////////// Khach Huy /////////////////////////////////////////////////////////////////////////////
        [HttpGet]
        public JsonResult LoadDataKhachHuy(string tungay, string denngay, string cn, string khoi, int page, int pageSize)
        {
            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            if (cn == null)
            {
                cn = "";
            }

            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            int totalRow = 0;

            var listAccount = _thongkeService.KhachHuyEntities(tungay, denngay, cn, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            //var responseData = Mapper.Map<IEnumerable<doanhthuSaleQuay>, IEnumerable<doanhthuSaleQuayViewModel>>(listAccount);

            return Json(new
            {
                data = listAccount,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ViewResult KhachHuy(string tungay, string denngay, string chinhanh, string khoi)//(string tungay,string denngay, string daily)
        {
            chinhanh = String.IsNullOrEmpty(chinhanh) ? Session["chinhanh"].ToString() : chinhanh;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//stt
            xlSheet.Column(2).Width = 30;// ten khach
            xlSheet.Column(3).Width = 30;// sgtcode
            xlSheet.Column(4).Width = 10;// Vetourid
            xlSheet.Column(5).Width = 40;// Tuyến tq
            
            xlSheet.Column(6).Width = 30;// Bắt đầu
            xlSheet.Column(7).Width = 30;// Kết thúc
            xlSheet.Column(8).Width = 20;// Giá tour
            xlSheet.Column(9).Width = 40;// Người hủy vé
            xlSheet.Column(10).Width = 30;// Đại lý hủy vé
            xlSheet.Column(11).Width = 10;// Chi nhánh
            xlSheet.Column(12).Width = 30;// Ngày hủy vé

            xlSheet.Cells[2, 1].Value = "THỐNG KÊ HỦY ĐOÀN " + khoi + " " + chinhanh;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 12].Merge = true;
            setCenterAligment(2, 1, 2, 12, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 12].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 12, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Tên khách";
            xlSheet.Cells[5, 3].Value = "Sgt code";
            xlSheet.Cells[5, 4].Value = "Vetourid";
            xlSheet.Cells[5, 5].Value = "Tuyến tq";
            
            xlSheet.Cells[5, 6].Value = "Bắt đầu";
            xlSheet.Cells[5, 7].Value = "Kết thúc";
            xlSheet.Cells[5, 8].Value = "Giá tour";
            xlSheet.Cells[5, 9].Value = "Người hủy vé";
            xlSheet.Cells[5, 10].Value = "Đại lý hủy vé";
            xlSheet.Cells[5, 11].Value = "Chi nhánh";            
            xlSheet.Cells[5, 12].Value = "Ngày hủy vé";

            xlSheet.Cells[5, 1, 5, 12].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table  
            int dong = 5;


            DataTable dt = _thongkeService.KhachHuyEntitiesToExcel(tungay, denngay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 8].Formula = "SUM(H6:D" + (6 + dt.Rows.Count - 1) + ")";

            setBorder(5, 1, 5 + dt.Rows.Count, 12, xlSheet);
            setFontBold(5, 1, 5, 12, 11, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 12, 11, xlSheet);
            // canh giua cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);
            // canh giua code chinhanh
            setCenterAligment(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 8, 6 + dt.Rows.Count, 8, xlSheet);
            // định dạng số cot tong cong
            NumberFormat(dong, 8, dong, 8, xlSheet);
            setBorder(dong, 8, dong, 8, xlSheet);
            setFontBold(dong, 8, dong, 8, 12, xlSheet);
            // DateFormat
            DateFormat(6, 6, 6 + dt.Rows.Count, 6, xlSheet);
            DateFormat(6, 7, 6 + dt.Rows.Count, 7, xlSheet);
            DateFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);
            
            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuay_" + Session["daily"].ToString() + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DSKhachHuy_" + khoi + " " + chinhanh + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            return View();
        }

        ////////////////////////////KDONLINE/////////////////////////////////////////////////////////////////////////////
        [HttpGet]
        public JsonResult LoadDataKinhDoanhOnline(string tungay, string denngay, string khoi, int page, int pageSize)
        {
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            int totalRow = 0;

            var listAccount = _thongkeService.KinhDoanhOnlineEntities(tungay, denngay, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<thongkeweb>, IEnumerable<thongkewebViewModel>>(listAccount);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult LoadDataKinhDoanhOnline_ngaydi(string tungay, string denngay, string khoi, int page, int pageSize)
        {
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;

            int totalRow = 0;

            var listAccount = _thongkeService.KinhDoanhOnlineEntities_ngaydi(tungay, denngay, khoi, page, pageSize, out totalRow);
            //var query = listuser.OrderBy(x => x.tenhd);
            var responseData = Mapper.Map<IEnumerable<thongkeweb>, IEnumerable<thongkewebViewModel>>(listAccount);

            return Json(new
            {
                data = responseData,
                total = totalRow,
                status = true
            }, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public ActionResult LoadDataKinhDoanhOnlineChitietToExcel(string tungay, string denngay, string chinhanh, string khoi)
        {

            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;//SGTCODE
            xlSheet.Column(3).Width = 35;// TUYEN TQ
            xlSheet.Column(4).Width = 15;// NGAY DI
            xlSheet.Column(5).Width = 15;// NGAY VE
            xlSheet.Column(6).Width = 30;// TEN KHACH
            xlSheet.Column(7).Width = 15;//  SERIAL
            xlSheet.Column(8).Width = 15;//  HUY VE
            xlSheet.Column(9).Width = 10;//  SÔ KHÁCH
            xlSheet.Column(10).Width = 15;//  DOANH SO
            xlSheet.Column(11).Width = 30;//  sale
            xlSheet.Column(12).Width = 30;//  DAI LY 
            xlSheet.Column(13).Width = 20;//  KENH GD


            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU ONLINE THEO NGÀY BÁN " + chinhanh;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 12].Merge = true;
            setCenterAligment(2, 1, 2, 12, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 13].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 13, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Sgt Code";
            xlSheet.Cells[5, 3].Value = "Hành trình ";
            xlSheet.Cells[5, 4].Value = "Ngày đi";
            xlSheet.Cells[5, 5].Value = "Ngày về";
            xlSheet.Cells[5, 6].Value = "Tên khách";
            xlSheet.Cells[5, 7].Value = "Serial";
            xlSheet.Cells[5, 8].Value = "Huỷ vé";
            xlSheet.Cells[5, 9].Value = "Số khách";
            xlSheet.Cells[5, 10].Value = "Doanh số";
            xlSheet.Cells[5, 11].Value = "Nhân viên";
            xlSheet.Cells[5, 12].Value = "Đại lý";
            xlSheet.Cells[5, 13].Value = "Kênh GD";
            xlSheet.Cells[5, 1, 5, 13].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.ThongkeWebchitiet(tungay, denngay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;

            xlSheet.Cells[dong, 8].Value = "TC";
            xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 10].Formula = "SUM(J6:J" + (6 + dt.Rows.Count - 1) + ")";
            // định dạng số
            NumberFormat(6, 10, 6 + dt.Rows.Count, 10, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 13, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 13, xlSheet);
            setFontBold(5, 1, 5, 10, 13, xlSheet);

            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot so khach
            setCenterAligment(6, 9, 6 + dt.Rows.Count, 9, xlSheet);

            setBorder(dong, 8, dong, 10, xlSheet);
            setFontBold(dong, 8, dong, 10, 12, xlSheet);
            // canh giưa cot ngay di va ngày ve
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            DateFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuKinhDoanhOnline" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            //return RedirectToAction("LoadDataQuayTheoNgayDiChitietToExcel");
            return View();
        }

        [HttpGet]
        public ActionResult LoadDataKinhDoanhOnline_NgayDi_ToExcel(string tungay, string denngay, string khoi)
        {
            
            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;//SGTCODE
            xlSheet.Column(3).Width = 35;// TUYEN TQ
            xlSheet.Column(4).Width = 15;// NGAY DI
            xlSheet.Column(5).Width = 15;// NGAY VE
            xlSheet.Column(6).Width = 30;// TEN KHACH
            xlSheet.Column(7).Width = 15;//  SERIAL
            xlSheet.Column(8).Width = 15;//  HUY VE
            xlSheet.Column(9).Width = 10;//  SÔ KHÁCH
            xlSheet.Column(10).Width = 15;//  DOANH SO
            xlSheet.Column(11).Width = 30;//  sale
            xlSheet.Column(12).Width = 30;//  DAI LY 
            xlSheet.Column(13).Width = 20;//  KENH GD


            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU ONLINE THEO NGÀY DI ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 12].Merge = true;
            setCenterAligment(2, 1, 2, 12, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 13].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 13, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Sgt Code";
            xlSheet.Cells[5, 3].Value = "Hành trình ";
            xlSheet.Cells[5, 4].Value = "Ngày đi";
            xlSheet.Cells[5, 5].Value = "Ngày về";
            xlSheet.Cells[5, 6].Value = "Tên khách";
            xlSheet.Cells[5, 7].Value = "Serial";
            xlSheet.Cells[5, 8].Value = "Huỷ vé";
            xlSheet.Cells[5, 9].Value = "Số khách";
            xlSheet.Cells[5, 10].Value = "Doanh số";
            xlSheet.Cells[5, 11].Value = "Nhân viên";
            xlSheet.Cells[5, 12].Value = "Đại lý";
            xlSheet.Cells[5, 13].Value = "Kênh GD";
            xlSheet.Cells[5, 1, 5, 13].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.ThongkeWebchitiet_ngaydi(tungay, denngay, "", khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;

            xlSheet.Cells[dong, 8].Value = "TC";
            xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 10].Formula = "SUM(J6:J" + (6 + dt.Rows.Count - 1) + ")";
            // định dạng số
            NumberFormat(6, 10, 6 + dt.Rows.Count, 10, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 13, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 13, xlSheet);
            setFontBold(5, 1, 5, 10, 13, xlSheet);

            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot so khach
            setCenterAligment(6, 9, 6 + dt.Rows.Count, 9, xlSheet);

            setBorder(dong, 8, dong, 10, xlSheet);
            setFontBold(dong, 8, dong, 10, 12, xlSheet);
            // canh giưa cot ngay di va ngày ve
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            DateFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuKinhDoanhOnline" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            //return RedirectToAction("LoadDataQuayTheoNgayDiChitietToExcel");
            return View();
        }

        [HttpGet]
        public ActionResult LoadDataKinhDoanhOnlineChitiet_ngaydiToExcel(string tungay, string denngay, string chinhanh, string khoi)
        {

            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;//SGTCODE
            xlSheet.Column(3).Width = 35;// TUYEN TQ
            xlSheet.Column(4).Width = 15;// NGAY DI
            xlSheet.Column(5).Width = 15;// NGAY VE
            xlSheet.Column(6).Width = 30;// TEN KHACH
            xlSheet.Column(7).Width = 15;//  SERIAL
            xlSheet.Column(8).Width = 15;//  HUY VE
            xlSheet.Column(9).Width = 10;//  SO KHACH
            xlSheet.Column(10).Width = 15;//  DOANH SO
            xlSheet.Column(11).Width = 30;//  sale
            xlSheet.Column(12).Width = 30;//  DAI LY 
            xlSheet.Column(13).Width = 20;//  KENH GD


            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU ONLINE THEO NGÀY ĐI TOUR " + chinhanh;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 13].Merge = true;
            setCenterAligment(2, 1, 2, 13, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 13].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 13, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Sgt Code";
            xlSheet.Cells[5, 3].Value = "Hành trình ";
            xlSheet.Cells[5, 4].Value = "Ngày đi";
            xlSheet.Cells[5, 5].Value = "Ngày về";
            xlSheet.Cells[5, 6].Value = "Tên khách";
            xlSheet.Cells[5, 7].Value = "Serial";
            xlSheet.Cells[5, 8].Value = "Huỷ vé";
            xlSheet.Cells[5, 9].Value = "Số khách";
            xlSheet.Cells[5, 10].Value = "Doanh số";
            xlSheet.Cells[5, 11].Value = "Nhân viên";
            xlSheet.Cells[5, 12].Value = "Đại lý";
            xlSheet.Cells[5, 13].Value = "Kênh GD";
            xlSheet.Cells[5, 1, 5, 13].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.ThongkeWebchitiet_ngaydi(tungay, denngay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return View();
            }

            dong++;

            xlSheet.Cells[dong, 8].Value = "TC";
            xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (6 + dt.Rows.Count - 1) + ")";
            xlSheet.Cells[dong, 10].Formula = "SUM(J6:J" + (6 + dt.Rows.Count - 1) + ")";
            // định dạng số
            NumberFormat(6, 10, 6 + dt.Rows.Count, 10, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 13, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 13, xlSheet);
            setFontBold(5, 1, 5, 10, 13, xlSheet);

            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot so khach
            setCenterAligment(6, 9, 6 + dt.Rows.Count, 9, xlSheet);

            setBorder(dong, 8, dong, 10, xlSheet);
            setFontBold(dong, 8, dong, 10, 12, xlSheet);
            
            // canh giưa cot ngay di va ngày ve
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 5, xlSheet);
            DateFormat(6, 4, 6 + dt.Rows.Count, 5, xlSheet);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuKinhDoanhOnline" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            //return RedirectToAction("LoadDataQuayTheoNgayDiChitietToExcel");
            return View();
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////


        [HttpGet]
        public ActionResult LoadDataQuayTheoNgayDiChitietToExcel(string quay, string chinhanh, string tungay, string denngay, string khoi)
        {

            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 10;//STT
            xlSheet.Column(3).Width = 25;// SGTCODE
            xlSheet.Column(4).Width = 15;// serial
            xlSheet.Column(5).Width = 30;// ten khach
            xlSheet.Column(6).Width = 40;// tuyen tq
            xlSheet.Column(7).Width = 15;//  ngay di
            xlSheet.Column(8).Width = 15;//  ngay ve
            xlSheet.Column(9).Width = 15;//  gia tour
            xlSheet.Column(10).Width = 30;//  sale


            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY ĐI QUẦY " + quay;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 10].Merge = true;
            setCenterAligment(2, 1, 2, 10, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 10].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 10, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Code CN";
            xlSheet.Cells[5, 3].Value = "Sgt Code ";
            xlSheet.Cells[5, 4].Value = "Serial";
            xlSheet.Cells[5, 5].Value = "Tên khách";
            xlSheet.Cells[5, 6].Value = "Hành trình";
            xlSheet.Cells[5, 7].Value = "Ngày đi";
            xlSheet.Cells[5, 8].Value = "Ngày về";
            xlSheet.Cells[5, 9].Value = "Doanh số";
            xlSheet.Cells[5, 10].Value = "Nhân viên";
            xlSheet.Cells[5, 1, 5, 10].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            DataTable dt = _thongkeService.doanhthuQuayTheoNgayDiChitiet(quay, chinhanh, tungay, denngay, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return PartialView(dt);
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 8].Value = "TC";
            xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số
            NumberFormat(dong, 8, dong, 8, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 11, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 10, xlSheet);
            setFontBold(5, 1, 5, 10, 12, xlSheet);

            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 2, xlSheet);

            setBorder(dong, 8, dong, 9, xlSheet);
            setFontBold(dong, 8, dong, 9, 12, xlSheet);
            // canh giưa cot ngay di va ngày ve
            setCenterAligment(6, 7, 6 + dt.Rows.Count, 8, xlSheet);
            // dinh dạng number cot gia ve
            NumberFormat(6, 9, 6 + dt.Rows.Count, 9, xlSheet);

            // xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuayChitiet" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            //return RedirectToAction("LoadDataQuayTheoNgayDiChitietToExcel");
            return PartialView(dt);
        }

        public ActionResult LoadDataQuayTheoNgayBanChitietToExcel(string tungay, string denngay, string quay, string chinhanh, string khoi)
        {

            //cn = String.IsNullOrEmpty(cn) ? Session["chinhanh"].ToString() : cn;
            khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
            string fromTo = "";
            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 10;//Code CN
            xlSheet.Column(3).Width = 25;// SGTCODE
            xlSheet.Column(4).Width = 15;// serial
            xlSheet.Column(5).Width = 30;// ten khach
            xlSheet.Column(6).Width = 40;// tuyen tq
            xlSheet.Column(7).Width = 15;//  ngay di
            xlSheet.Column(8).Width = 15;//  ngay ve
            xlSheet.Column(9).Width = 15;//  gia tour
            xlSheet.Column(10).Width = 30;//  sale


            xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY BÁN QUẦY " + quay;
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 10].Merge = true;
            setCenterAligment(2, 1, 2, 10, xlSheet);
            // dinh dang tu ngay den ngay
            if (tungay == denngay)
            {
                fromTo = "Ngày: " + tungay;
            }
            else
            {
                fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            }
            xlSheet.Cells[3, 1].Value = fromTo;
            xlSheet.Cells[3, 1, 3, 10].Merge = true;
            xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            setCenterAligment(3, 1, 3, 10, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Code CN";
            xlSheet.Cells[5, 3].Value = "Sgt Code ";
            xlSheet.Cells[5, 4].Value = "Serial";
            xlSheet.Cells[5, 5].Value = "Tên khách";
            xlSheet.Cells[5, 6].Value = "Hành trình";
            xlSheet.Cells[5, 7].Value = "Ngày đi";
            xlSheet.Cells[5, 8].Value = "Ngày về";
            xlSheet.Cells[5, 9].Value = "Doanh số";
            xlSheet.Cells[5, 10].Value = "Nhân viên";
            xlSheet.Cells[5, 1, 5, 10].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;

            //DataTable dt = _thongkeService.doanhthuQuayTheoNgayBanChitiet(quay, chinhanh, tungay, denngay, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());
            DataTable dt = _thongkeService.doanhthuQuayTheoNgayBanChitiet(tungay, denngay, quay, chinhanh, khoi);// Session["daily"].ToString(), Session["khoi"].ToString());

            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = 0;
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                SetAlert("No sale.", "warning");
                return PartialView(dt);
            }

            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";
            // Sum tổng tiền
            xlSheet.Cells[dong, 8].Value = "TC";
            xlSheet.Cells[dong, 9].Formula = "SUM(I6:I" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số
            NumberFormat(dong, 8, dong, 8, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 11, 12, xlSheet);
            setBorder(5, 1, 5 + dt.Rows.Count, 10, xlSheet);
            setFontBold(5, 1, 5, 10, 12, xlSheet);

            // canh giưa cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 2, xlSheet);

            setBorder(dong, 8, dong, 9, xlSheet);
            setFontBold(dong, 8, dong, 9, 12, xlSheet);
            // canh giưa cot ngay di va ngày ve
            setCenterAligment(6, 7, 6 + dt.Rows.Count, 8, xlSheet);
            // dinh dạng number cot gia ve
            NumberFormat(6, 9, 6 + dt.Rows.Count, 9, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuQuayChitiet" + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
            Response.BinaryWrite(ExcelApp.GetAsByteArray());
            Response.End();

            //return RedirectToAction("LoadDataQuayTheoNgayDiChitietToExcel");
            return PartialView(dt);
        }

        public ViewResult LoadDataSaleTheoNgayBanChitietToExcel(string tungay, string denngay, string nhanvien, string chinhanh, string khoi)
        {
            try
            {
                nhanvien = convertToUnSign3(nhanvien);
                khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
                string fromTo = "";
                ExcelPackage ExcelApp = new ExcelPackage();
                ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("DoanhthuSale");
                // Định dạng chiều dài cho cột
                xlSheet.Column(1).Width = 10;//stt
                xlSheet.Column(2).Width = 10;// chi nhanh
                xlSheet.Column(3).Width = 25;// code
                xlSheet.Column(4).Width = 25;// tuyen tham quan
                xlSheet.Column(5).Width = 40;// ten khach
                xlSheet.Column(6).Width = 10;// so khach
                xlSheet.Column(7).Width = 20;//doanhthu
                xlSheet.Column(8).Width = 20;//thuc thu
                xlSheet.Column(9).Width = 35;//sales

                xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY BÁN SALE " + nhanvien;
                xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
                xlSheet.Cells[2, 1, 2, 9].Merge = true;
                if (tungay == denngay)
                {
                    fromTo = "Ngày: " + tungay;
                }
                else
                {
                    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
                }
                xlSheet.Cells[3, 1].Value = fromTo;
                xlSheet.Cells[3, 1, 3, 9].Merge = true;
                xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
                setCenterAligment(2, 1, 3, 9, xlSheet);

                // Tạo header
                xlSheet.Cells[5, 1].Value = "STT";
                xlSheet.Cells[5, 2].Value = "Code CN";
                xlSheet.Cells[5, 3].Value = "Code Đoàn";
                xlSheet.Cells[5, 4].Value = "Tuyến tham quan";
                xlSheet.Cells[5, 5].Value = "Tên khách";
                xlSheet.Cells[5, 6].Value = "Số khách";
                xlSheet.Cells[5, 7].Value = "Tổng tiền";
                xlSheet.Cells[5, 8].Value = "Doanh số";
                xlSheet.Cells[5, 9].Value = "Sales";

                xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

                int dong = 5;
                DataTable dt = _thongkeService.doanhthuSaleTheoNgayBanChitiet(tungay, denngay, nhanvien, chinhanh, khoi);// Session["fullName"].ToString());

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dong++;
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                            {
                                xlSheet.Cells[dong, j + 1].Value = 0;
                            }
                            else
                            {
                                xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                            }
                        }
                    }
                }
                else
                {
                    SetAlert("No sale.", "warning");
                    return View();
                }
                dong++;
                // Merger cot 4,5 ghi tổng tiền
                //setRightAligment(dong, 4, dong, 5, xlSheet);
                //xlSheet.Cells[dong, 4, dong, 5].Merge = true;
                //xlSheet.Cells[dong, 4].Value = "Tổng tiền: ";

                //// Sum tổng tiền
                xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (6 + dt.Rows.Count - 1) + ")";
                //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";
                //// định dạng số
                NumberFormat(dong, 6, dong, 6, xlSheet);
                setBorder(5, 1, 5 + dt.Rows.Count, 9, xlSheet);
                setFontBold(5, 1, 5, 9, 12, xlSheet);
                setFontSize(6, 1, 6 + dt.Rows.Count, 9, 12, xlSheet);
                NumberFormat(6, 7, 6 + dt.Rows.Count, 8, xlSheet);
                setCenterAligment(6, 1, 6 + dt.Rows.Count, 3, xlSheet);
                setCenterAligment(6, 6, 6 + dt.Rows.Count, 6, xlSheet);
                xlSheet.View.FreezePanes(6, 20);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuSale_" + nhanvien + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                Response.BinaryWrite(ExcelApp.GetAsByteArray());
                Response.End();

            }
            catch
            {
                throw;
            }
            return View();
        }

        public ViewResult LoadDataSaleTheoNgayDiChitietToExcel(string tungay, string denngay, string nhanvien, string chinhanh, string khoi)
        {
            try
            {
                nhanvien = convertToUnSign3(nhanvien);
                khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
                string fromTo = "";
                ExcelPackage ExcelApp = new ExcelPackage();
                ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("DoanhthuSale");
                // Định dạng chiều dài cho cột
                xlSheet.Column(1).Width = 10;//stt
                xlSheet.Column(2).Width = 10;// chi nhanh
                xlSheet.Column(3).Width = 25;// sgtcode
                xlSheet.Column(4).Width = 25;// tuyen tham quan
                xlSheet.Column(5).Width = 40;// ten khach
                xlSheet.Column(6).Width = 10;// so khach
                xlSheet.Column(7).Width = 20;//doanhthu
                xlSheet.Column(8).Width = 20;//thuc thu
                xlSheet.Column(9).Width = 35;//sales

                xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU THEO NGÀY ĐI SALE " + nhanvien;
                xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
                xlSheet.Cells[2, 1, 2, 8].Merge = true;
                if (tungay == denngay)
                {
                    fromTo = "Ngày: " + tungay;
                }
                else
                {
                    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
                }
                xlSheet.Cells[3, 1].Value = fromTo;
                xlSheet.Cells[3, 1, 3, 9].Merge = true;
                xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
                setCenterAligment(2, 1, 3, 9, xlSheet);

                // Tạo header
                xlSheet.Cells[5, 1].Value = "STT";
                xlSheet.Cells[5, 2].Value = "Code CN";
                xlSheet.Cells[5, 3].Value = "Code Đoàn";
                xlSheet.Cells[5, 4].Value = "Tuyến tham quan";
                xlSheet.Cells[5, 5].Value = "Tên khách";
                xlSheet.Cells[5, 6].Value = "Số khách";
                xlSheet.Cells[5, 7].Value = "Tổng tiền";
                xlSheet.Cells[5, 8].Value = "Doanh số";
                xlSheet.Cells[5, 9].Value = "Sales";

                xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

                int dong = 5;
                DataTable dt = _thongkeService.doanhthuSaleTheoNgayDiChitiet(tungay, denngay, nhanvien, chinhanh, khoi);// Session["fullName"].ToString());

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dong++;
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                            {
                                xlSheet.Cells[dong, j + 1].Value = 0;
                            }
                            else
                            {
                                xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                            }
                        }
                    }
                }
                else
                {
                    SetAlert("No sale.", "warning");
                    return View();
                }
                dong++;
                // Merger cot 4,5 ghi tổng tiền
                //setRightAligment(dong, 4, dong, 5, xlSheet);
                //xlSheet.Cells[dong, 4, dong, 5].Merge = true;
                //xlSheet.Cells[dong, 4].Value = "Tổng tiền: ";

                //// Sum tổng tiền
                xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (6 + dt.Rows.Count - 1) + ")";
                //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";
                //// định dạng số
                NumberFormat(dong, 7, dong, 8, xlSheet);
                setBorder(5, 1, 5 + dt.Rows.Count, 9, xlSheet);
                setFontBold(5, 1, 5, 9, 12, xlSheet);
                setFontSize(6, 1, 6 + dt.Rows.Count, 9, 12, xlSheet);
                NumberFormat(6, 7, 6 + dt.Rows.Count, 8, xlSheet);
                setCenterAligment(6, 1, 6 + dt.Rows.Count, 3, xlSheet);
                setCenterAligment(6, 6, 6 + dt.Rows.Count, 6, xlSheet);
                xlSheet.View.FreezePanes(6, 20);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuSale_" + nhanvien + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                Response.BinaryWrite(ExcelApp.GetAsByteArray());
                Response.End();

            }
            catch
            {
                throw;
            }
            return View();
        }
        public ViewResult LoadDataDoanTheoNgayDiChitietToExcel(string sgtcode, string khoi)
        {
            try
            {
                khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
                string fromTo = "";
                DataTable d = _thongkeService.getTourbySgtcode(sgtcode, khoi);
                string tuyentq = d.Rows[0]["tuyentq"].ToString();
                string ngay = "ĐOÀN: " + sgtcode + " bắt đầu: " + Convert.ToDateTime(d.Rows[0]["batdau"]).ToString("dd/MM/yyyy HH:mm") + " kết thúc: " + Convert.ToDateTime(d.Rows[0]["ketthuc"]).ToString("dd/MM/yyyy HH:mm");
                ExcelPackage ExcelApp = new ExcelPackage();
                ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("DoanhthuDoan");
                // Định dạng chiều dài cho cột
                xlSheet.Column(1).Width = 7;//stt
                xlSheet.Column(2).Width = 20;// Serial
                xlSheet.Column(3).Width = 40;// Ten khach
                xlSheet.Column(4).Width = 45;// Dia chi
                xlSheet.Column(5).Width = 30;// Diem don
                xlSheet.Column(6).Width =10;// Gia ve
                xlSheet.Column(7).Width = 10;//Thuc thu
                xlSheet.Column(8).Width = 10;//Cong no
                xlSheet.Column(9).Width = 45;//Ghi chu

                xlSheet.Cells[2, 1].Value = tuyentq;// "BÁO CÁO DOANH THU THEO NGÀY ĐI " + sgtcode;
                xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
                xlSheet.Cells[2, 1, 2, 8].Merge = true;

                xlSheet.Cells[3, 1].Value = ngay;
                xlSheet.Cells[3, 1, 3, 8].Merge = true;
                xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Regular));
                setCenterAligment(2, 1, 3, 8, xlSheet);

                // Tạo header
                xlSheet.Cells[5, 1].Value = "STT";
                xlSheet.Cells[5, 2].Value = "Serial";
                xlSheet.Cells[5, 3].Value = "Tên khách";
                xlSheet.Cells[5, 4].Value = "Địa chỉ - Tel";
                xlSheet.Cells[5, 5].Value = "Điểm đón";
                xlSheet.Cells[5, 6].Value = "Giá vé";
                xlSheet.Cells[5, 7].Value = "Thực thu";
                xlSheet.Cells[5, 8].Value = "Công nợ";
                xlSheet.Cells[5, 9].Value = "Ghi chú";
               
                xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 10, FontStyle.Bold));
                setBorder(5, 1, 5 , 9, xlSheet);
                int dong = 5;
                int giongnhau = 0;
                DataTable dt = _thongkeService.doanhthuDoanTheoNgayDiChitiet(sgtcode, khoi);// Session["fullName"].ToString());

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dong++;
                        for (int j = 2; j < dt.Columns.Count; j++)
                        {
                            if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                            {
                                xlSheet.Cells[dong, j - 1].Value = " ";
                            }
                            else
                            {
                                xlSheet.Cells[dong, j - 1].Value = dt.Rows[i][j];
                            }
                        }
                        if (i > 0 && dt.Rows[i][1].ToString() == dt.Rows[i - 1][1].ToString())
                        {
                            giongnhau++;
                            
                        }
                        else
                        {
                            giongnhau = 0;
                        }
                        if (giongnhau > 0)
                        {
                            mergercell(dong - giongnhau, 2, dong, 2, xlSheet);
                            mergercell(dong - giongnhau, 5, dong, 5, xlSheet);
                            numberMergercell(dong - giongnhau, 6, dong, 6, xlSheet);
                            numberMergercell(dong - giongnhau, 7, dong, 7, xlSheet);
                            numberMergercell(dong - giongnhau, 8, dong, 8, xlSheet);
                            mergercell(dong - giongnhau, 9, dong, 9, xlSheet);
                            setBorderAround(dong - giongnhau, 1, dong, 1, xlSheet);
                            setBorderAround(dong - giongnhau, 2, dong, 2, xlSheet);
                            setBorderAround(dong - giongnhau, 3, dong, 3, xlSheet);
                            setBorderAround(dong - giongnhau, 4, dong, 4, xlSheet);
                            setBorderAround(dong - giongnhau, 5, dong, 5, xlSheet);
                            setBorderAround(dong - giongnhau, 6, dong, 6, xlSheet);
                            setBorderAround(dong - giongnhau, 7, dong, 7, xlSheet);
                            setBorderAround(dong - giongnhau, 8, dong, 8, xlSheet);
                            setBorderAround(dong - giongnhau, 9, dong, 9, xlSheet);
                        }
                        else
                        {
                            wrapText(dong, 2, dong, 2, xlSheet);
                            wrapText(dong, 5, dong, 9, xlSheet);
                            wrapText(dong, 9, dong, 9, xlSheet);
                            setBorderAround(dong, 1, dong, 1, xlSheet);
                            setBorderAround(dong , 2, dong, 2, xlSheet);
                            setBorderAround(dong, 3, dong, 3, xlSheet);
                            setBorderAround(dong , 4, dong, 4, xlSheet);
                            setBorderAround(dong , 5, dong, 5, xlSheet);
                            setBorderAround(dong , 6, dong, 6, xlSheet);
                            setBorderAround(dong , 7, dong, 7, xlSheet);
                            setBorderAround(dong , 8, dong, 8, xlSheet);
                            setBorderAround(dong , 9, dong, 9, xlSheet);
                        }
                        
                    }
                    
                }
                else
                {
                    SetAlert("No sale.", "warning");
                    return View();
                }
                dong++;
                // set border
              //  setBorder(5, 1, 5 + dt.Rows.Count, 9, xlSheet);
                setFontSize(6, 1, 6 + dt.Rows.Count+1, 9,9, xlSheet);
                setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);
                wrapText(6, 6, 6 + dt.Rows.Count + 1, 8, xlSheet);

                //// Sum tổng tiền
                xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
                NumberFormat(6, 6, 6 + dt.Rows.Count +1, 8, xlSheet);
                
              
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuDoan_" + sgtcode + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                Response.BinaryWrite(ExcelApp.GetAsByteArray());
                Response.End();

            }
            catch
            {
                throw;
            }
            return View();
        }

        public void LoadDataSaleTheoTuyentqChitietToExcel(string tungay, string denngay, string nhanvien, string tuyentq, string khoi)
        {
            try
            {
                nhanvien = convertToUnSign3(nhanvien);
                khoi = String.IsNullOrEmpty(khoi) ? Session["khoi"].ToString() : khoi;
                string fromTo = "";
                ExcelPackage ExcelApp = new ExcelPackage();
                ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("DoanhthuSale");
                // Định dạng chiều dài cho cột
                xlSheet.Column(1).Width = 10;//stt
                xlSheet.Column(2).Width = 10;// chi nhanh
                xlSheet.Column(3).Width = 25;// sgtcode
                xlSheet.Column(4).Width = 25;// tuyen tham quan
                xlSheet.Column(5).Width = 40;// ten khach
                xlSheet.Column(6).Width = 10;// so khach
                xlSheet.Column(7).Width = 20;//doanhthu
                xlSheet.Column(8).Width = 20;//thuc thu
                xlSheet.Column(9).Width = 35;//sales

                xlSheet.Cells[2, 1].Value = "BÁO CÁO DOANH THU SALE THEO TUYEN " + tuyentq.ToUpper();
                xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
                xlSheet.Cells[2, 1, 2, 8].Merge = true;
                if (tungay == denngay)
                {
                    fromTo = "Ngày: " + tungay;
                }
                else
                {
                    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
                }
                xlSheet.Cells[3, 1].Value = fromTo;
                xlSheet.Cells[3, 1, 3, 9].Merge = true;
                xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
                setCenterAligment(2, 1, 3, 9, xlSheet);

                // Tạo header
                xlSheet.Cells[5, 1].Value = "STT";
                xlSheet.Cells[5, 2].Value = "Code CN";
                xlSheet.Cells[5, 3].Value = "Code Đoàn";
                xlSheet.Cells[5, 4].Value = "Tuyến tham quan";
                xlSheet.Cells[5, 5].Value = "Tên khách";
                xlSheet.Cells[5, 6].Value = "Số khách";
                xlSheet.Cells[5, 7].Value = "Tổng tiền";
                xlSheet.Cells[5, 8].Value = "Doanh số";
                xlSheet.Cells[5, 9].Value = "Sales";

                xlSheet.Cells[5, 1, 5, 9].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

                int dong = 5;
                DataTable dt = _thongkeService.DoanhThuSaleChitietTuyentq(tungay, denngay, nhanvien, tuyentq, khoi);// Session["fullName"].ToString());

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dong++;
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                            {
                                xlSheet.Cells[dong, j + 1].Value = 0;
                            }
                            else
                            {
                                xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                            }
                        }
                    }
                }
                else
                {
                    SetAlert("No sale.", "warning");
                    // return View();
                }
                dong++;
                // Merger cot 4,5 ghi tổng tiền
                //setRightAligment(dong, 4, dong, 5, xlSheet);
                //xlSheet.Cells[dong, 4, dong, 5].Merge = true;
                //xlSheet.Cells[dong, 4].Value = "Tổng tiền: ";

                //// Sum tổng tiền
                xlSheet.Cells[dong, 8].Formula = "SUM(H6:H" + (6 + dt.Rows.Count - 1) + ")";
                //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";
                //// định dạng số
                NumberFormat(dong, 7, dong, 8, xlSheet);
                setBorder(5, 1, 5 + dt.Rows.Count, 9, xlSheet);
                setFontBold(5, 1, 5, 9, 12, xlSheet);
                setFontSize(6, 1, 6 + dt.Rows.Count, 9, 12, xlSheet);
                NumberFormat(6, 7, 6 + dt.Rows.Count, 8, xlSheet);
                setCenterAligment(6, 1, 6 + dt.Rows.Count, 3, xlSheet);
                setCenterAligment(6, 6, 6 + dt.Rows.Count, 6, xlSheet);
                xlSheet.View.FreezePanes(6, 20);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=" + "DoanhThuTheoTuyentqSale_" + nhanvien + "_" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                Response.BinaryWrite(ExcelApp.GetAsByteArray());
                Response.End();

            }
            catch
            {
                throw;
            }
            // return View();
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////


        private static void NumberFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                range.Style.Numberformat.Format = "#,#0";
            }
        }
        private static void DateFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Numberformat.Format = "dd/MM/yyyy";
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void DateTimeFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void setRightAligment(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
        }
        private static void setCenterAligment(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void setFontSize(int fromRow, int fromColumn, int toRow, int toColumn, int fontSize, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Font.SetFromFont(new Font("Times New Roman", fontSize, FontStyle.Regular));
            }
        }
        private static void setFontBold(int fromRow, int fromColumn, int toRow, int toColumn, int fontSize, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Font.SetFromFont(new Font("Times New Roman", fontSize, FontStyle.Bold));
            }
        }
        private static void setBorder(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            }
        }
        private static void setBorderAround(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {               
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);              
            }
        }
        private static void PhantramFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Numberformat.Format = "0 %";
            }
        }
        private static void mergercell(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Merge = true;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;                              
                range.Style.WrapText = true;
            }
            
        }
        private static void numberMergercell(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                var a = sheet.Cells[fromRow, fromColumn].Value;
                range.Merge = true;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                range.Clear();
                sheet.Cells[fromRow, fromColumn].Value = a;
            }
        }
        private static void wrapText(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {               
                range.Style.WrapText = true;
            }

        }
        public static string convertToUnSign3(string s)
        {
            Regex regex = new Regex("\\p{IsCombiningDiacriticalMarks}+");
            string temp = s.Normalize(NormalizationForm.FormD);
            return regex.Replace(temp, String.Empty).Replace('\u0111', 'd').Replace('\u0110', 'D');
        }
    }
}