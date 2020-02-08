using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Infrastructure;
using ThongKe.Data.Models;
using ThongKe.Data.Models.EF;
using ThongKe.Data.Repositories;

namespace ThongKe.Service
{
    public interface IThongKeService
    {
        DataTable doanhthuSaleTheoQuay(string tungay, string denngay, string cn, string khoi);

        IEnumerable<doanhthuSaleQuay> doanhthuSaleTheoQuayEntities(string tungay, string denngay, string cn, string khoi, int page, int pageSize, out int totalRow);
        IEnumerable<doanhthuSaleTuyen> doanhthuSaleTheoTuyentqEntities(string tungay, string denngay, string tuyentq, string khoi, int page, int pageSize, out int totalRow);
        IEnumerable<thongkeweb> KinhDoanhOnlineEntities(string tungay,string denngay,string khoi,int page,int pageSize, out int totalRow);
        IEnumerable<thongkeweb> KinhDoanhOnlineEntities_ngaydi(string tungay, string denngay, string khoi, int page, int pageSize, out int totalRow);
        DataTable doanhthuSaleTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi);

        IEnumerable<doanhthuSaleQuay> doanhthuSaleTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);

        DataTable doanhthuQuayTheoNgayBan(string tungay, string denngay, string chinhanh, string khoi);

        IEnumerable<doanthuQuayNgayBan> doanhthuQuayTheoNgayBanEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);


        DataTable doanhthuQuayTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi);

        IEnumerable<doanthuQuayNgayBan> doanhthuQuayTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);

        DataTable doanhthuDoanTheoNgay(string tungay, string denngay, string chinhanh, string khoi);

        IEnumerable<doanhthuDoanNgayDi> doanhthuDoanTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);

        DataTable doanhthuTuyentqTheoNgay(string tungay, string denngay, string chinhanh, string khoi);

        DataTable doanhthuTuyentqTheoQuy(int quy, int nam, string cn, string khoi);

        IEnumerable<tuyentqNgaydi> doanhthuTuyentqTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);

        DataTable doanhthuKhachleHethong(string tungay, string denngay, string chinhanh, string khoi);
        IEnumerable<doanhthuToanhethong> doanhthuKhachLeHeThongEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow);

        DataTable doanhthuQuayTheoNgayDiChitiet(string quay, string chinhanh, string tungay, string denngay, string khoi);

        DataTable doanhthuQuayTheoNgayBanChitiet(string tungay, string denngay, string quay, string chinhanh, string khoi);

        DataTable doanhthuSaleTheoNgayBanChitiet(string tungay, string denngay, string nhanvien, string chinhanh, string khoi);
        DataTable doanhthuSaleTheoNgayDiChitiet(string tungay, string denngay, string nhanvien, string chinhanh, string khoi);

        DataTable ThongkeWebchitiet(string tungay, string denngay, string chinhanh, string khoi);
        DataTable doanhthuDoanTheoNgayDiChitiet(string sgtcode, string khoi);
        DataTable getTourbySgtcode(string sgtcode, string khoi);
        DataTable ThongkeWebchitiet_ngaydi(string tungay, string denngay, string chinhanh, string khoi);
        DataTable doanhthuSaleTheoTuyettq(string tungay, string denngay, string chinhanh, string khoi);
        DataTable DoanhThuSaleChitietTuyentq(string tungay, string denngay, string nhanvien, string tuyentq, string khoi);
        ///////////////////////////////////////////// Khach Huy ///////////////////////////////////////////////////////////////
        IEnumerable<KhachHuy> KhachHuyEntities(string tungay, string denngay, string cn, string khoi, int page, int pageSize, out int totalRow);
        DataTable KhachHuyEntitiesToExcel(string tungay, string denngay, string cn, string khoi);
        ////////////////////////////////// Tuyentq theo ngày di chi tiet ////////////////////////////////////////////
        DataTable doanhthuTuyentqTheoNgayDiChitiet(string tungay, string denngay, string chinhanh, string tuyentq, string khoi);
    }
    public class ThongKeService : IThongKeService
    {
        private IthongkeRepository _thongkeRepository;
        private IUnitOfWork _unitOfWork;

        public ThongKeService(IthongkeRepository thongkeRepository, IUnitOfWork unitOfWork)
        {
            _thongkeRepository = thongkeRepository;
            _unitOfWork = unitOfWork;
        }

        public DataTable doanhthuDoanTheoNgay(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuDoanTheoNgay(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public DataTable doanhthuKhachleHethong(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuKhachleHethong(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public DataTable doanhthuQuayTheoNgayBan(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuQuayTheoNgayBan(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public DataTable doanhthuQuayTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuQuayTheoNgayDi(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public IEnumerable<doanthuQuayNgayBan> doanhthuQuayTheoNgayBanEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuQuayTheoNgayBanEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable doanhthuSaleTheoNgayDi(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuSaleTheoNgayDi(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public IEnumerable<doanhthuSaleQuay> doanhthuSaleTheoNgayDiEntities(string tungay, string denngay, string cn, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuSaleTheoNgayDiEntities(tungay, denngay, cn, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable doanhthuSaleTheoQuay(string tungay, string denngay, string cn, string khoi)
        {
            var result = _thongkeRepository.doanhthuSaleTheoQuay(tungay, denngay, cn, khoi);
            return result;
        }

        public IEnumerable<doanhthuSaleQuay> doanhthuSaleTheoQuayEntities(string tungay, string denngay, string cn, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuSaleTheoQuayEntities(tungay, denngay, cn, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable doanhthuTuyentqTheoNgay(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuTuyentqTheoNgay(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public IEnumerable<doanthuQuayNgayBan> doanhthuQuayTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuQuayTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public IEnumerable<doanhthuDoanNgayDi> doanhthuDoanTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuDoanTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public IEnumerable<tuyentqNgaydi> doanhthuTuyentqTheoNgayDiEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuTuyentqTheoNgayDiEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public IEnumerable<doanhthuToanhethong> doanhthuKhachLeHeThongEntities(string tungay, string denngay, string chinhanh, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuKhachLeHeThongEntities(tungay, denngay, chinhanh, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable doanhthuQuayTheoNgayDiChitiet(string quay, string chinhanh, string tungay, string denngay, string khoi)
        {
            var result = _thongkeRepository.doanhthuQuayTheoNgayDiChitiet(quay, chinhanh, tungay, denngay, khoi);
            return result;
        }

        public DataTable doanhthuQuayTheoNgayBanChitiet(string tungay, string denngay, string quay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuQuayTheoNgayBanChitiet(tungay, denngay, quay, chinhanh, khoi);
            return result;
        }

        public DataTable doanhthuSaleTheoNgayBanChitiet(string tungay, string denngay, string nhanvien, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuSaleTheoNgayBanChitiet(tungay, denngay, nhanvien, chinhanh, khoi);
            return result;
        }

        public DataTable doanhthuSaleTheoNgayDiChitiet(string tungay, string denngay, string nhanvien, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuSaleTheoNgayDiChitiet(tungay, denngay, nhanvien, chinhanh, khoi);
            return result;
        }

        public IEnumerable<thongkeweb> KinhDoanhOnlineEntities(string tungay, string denngay, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.KinhDoanhOnlineEntities(tungay, denngay,  khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }
        public IEnumerable<thongkeweb> KinhDoanhOnlineEntities_ngaydi(string tungay, string denngay, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.KinhDoanhOnlineEntities_ngaydi(tungay, denngay, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }
        public DataTable ThongkeWebchitiet(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.ThongkeWebchitiet(tungay, denngay, chinhanh, khoi);
            return result;
        }
        public DataTable ThongkeWebchitiet_ngaydi(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.ThongkeWebchitiet_ngaydi(tungay, denngay, chinhanh, khoi);
            return result;
        }
        public DataTable doanhthuDoanTheoNgayDiChitiet(string sgtcode, string khoi)
        {
            var result = _thongkeRepository.doanhthuDoanTheoNgayDiChitiet(sgtcode, khoi);
            return result;
        }

        public DataTable getTourbySgtcode(string sgtcode, string khoi)
        {
            var result = _thongkeRepository.getTourbySgtcode(sgtcode, khoi);
            return result;
        }

        public DataTable doanhthuSaleTheoTuyettq(string tungay, string denngay, string chinhanh, string khoi)
        {
            var result = _thongkeRepository.doanhthuSaleTheoTuyettq(tungay, denngay, chinhanh, khoi);
            return result;
        }

        public IEnumerable<doanhthuSaleTuyen> doanhthuSaleTheoTuyentqEntities(string tungay, string denngay, string tuyentq, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.doanhthuSaleTheoTuyentqEntities(tungay, denngay, tuyentq, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable DoanhThuSaleChitietTuyentq(string tungay, string denngay, string nhanvien, string tuyentq, string khoi)
        {
            var listDanhThu = _thongkeRepository.DoanhThuSaleChitietTuyentq(tungay, denngay, nhanvien, tuyentq, khoi);
            return listDanhThu;
        }

        public DataTable doanhthuTuyentqTheoQuy(int quy, int nam, string cn, string khoi)
        {
            var result = _thongkeRepository.doanhthuTuyentqTheoQuy(quy, nam, cn, khoi);
            return result;
        }

        //////////////////////////////////////// Khach Huy /////////////////////////////////////////////////////
        public IEnumerable<KhachHuy> KhachHuyEntities(string tungay, string denngay, string cn, string khoi, int page, int pageSize, out int totalRow)
        {
            var listDanhThu = _thongkeRepository.KhachHuyEntities(tungay, denngay, cn, khoi, page, pageSize, out totalRow);
            return listDanhThu;
        }

        public DataTable KhachHuyEntitiesToExcel(string tungay, string denngay, string cn, string khoi)
        {
            var result = _thongkeRepository.KhachHuyEntitiesToExcel(tungay, denngay, cn, khoi);
            return result;
        }

        public DataTable doanhthuTuyentqTheoNgayDiChitiet(string tungay, string denngay, string chinhanh, string tuyentq, string khoi)
        {
            var result = _thongkeRepository.doanhthuTuyentqTheoNgayDiChitiet(tungay, denngay, chinhanh, tuyentq, khoi);
            return result;
        }
    }
}
