using System.Collections.Generic;
using System.Linq;
using ThongKe.Common.ViewModel;
using ThongKe.Data.Infrastructure;
using ThongKe.Data.Models.EF;
using ThongKe.Data.Repositories;

namespace ThongKe.Service
{
    public interface ICommonService
    {
        IEnumerable<string> GetAllChiNhanh();
        IEnumerable<string> GetAllChiNhanhByNhom(string nhom,string chinhanh);

        IEnumerable<dmdaily> GetDmdailyByChiNhanh(string chinhanh);
        IEnumerable<dmdaily> GetDmdailyByNhomChiNhanh(string nhom);

        IEnumerable<dmdaily> GetAllDmDaiLy();

        IEnumerable<ThongKeKhachViewModel> ThongKeSoKhachOB(string khoi);

        IEnumerable<ThongKeDoanhThuViewModel> ThongKeDoanhThuOB(string khoi);
        IEnumerable<string> GetAllTuyentqByKhoi(string khoi);
    }

    public class CommonService : ICommonService
    {
        private IchinhanhRepository _chinhanhRepository;
        private IdmdailyRepository _dmdailyRepository;
        private IthongkeRepository _thongkeRepository;
        private IUnitOfWork _unitOfWork;

        public CommonService(IchinhanhRepository chinhanhRepository, IdmdailyRepository dmdailyRepository, IthongkeRepository thongkeRepository, IUnitOfWork unitOfWork)
        {
            _chinhanhRepository = chinhanhRepository;
            _dmdailyRepository = dmdailyRepository;
            _thongkeRepository = thongkeRepository;
            _unitOfWork = unitOfWork;
        }

        public IEnumerable<string> GetAllChiNhanh()
        {
            return _chinhanhRepository.GetAll().Select(x=>x.chinhanh1).Distinct();
        }

        public IEnumerable<string> GetAllChiNhanhByNhom(string nhom,string chinhanh)
        {
            var result = new List<string>();
            if (nhom == "Admins")
            {
                result = _chinhanhRepository.GetAll().Select(x => x.chinhanh1).Distinct().ToList();
            }
            else if(nhom=="TNB"||nhom=="DNB"||nhom=="MT"||nhom=="MB")
            {
                result = _chinhanhRepository.GetMulti(x => x.nhom == nhom).Select(x => x.chinhanh1).Distinct().ToList();
                var count = result.Count();
            }
            else
            {
                result.Add(_chinhanhRepository.GetSingleByCondition(x => x.chinhanh1 == chinhanh).chinhanh1);
            }

            return result;
        }

        public IEnumerable<dmdaily> GetAllDmDaiLy()
        {
            return _dmdailyRepository.GetAll();
        }

        public IEnumerable<string> GetAllTuyentqByKhoi(string khoi)
        {
            var listTuyentq = _thongkeRepository.GetAllTuyentqByKhoi(khoi);
            return listTuyentq;
        }

        public IEnumerable<dmdaily> GetDmdailyByChiNhanh(string chinhanh)
        {
            return _dmdailyRepository.GetMulti(x => x.chinhanh == chinhanh);
        }

        public IEnumerable<dmdaily> GetDmdailyByNhomChiNhanh(string nhom)
        {
            var listDaily = new List<dmdaily>();
            switch (nhom)
            {
                //case "Admins":
                //    listDaily = _dmdailyRepository.GetAll().ToList();
                //    break;
                case "TNB":
                    var listChinhanhTNB = _chinhanhRepository.GetMulti(x => x.nhom == nhom).ToList();
                    foreach (var chinhanh in listChinhanhTNB)
                    {
                        var daily = _dmdailyRepository.GetSingleByCondition(x => x.chinhanh == chinhanh.chinhanh1);
                        listDaily.Add(daily);
                    }
                    break;
                case "DNB":
                    var listChinhanhDNB = _chinhanhRepository.GetMulti(x => x.nhom == nhom).ToList();
                    foreach (var chinhanh in listChinhanhDNB)
                    {
                        var daily = _dmdailyRepository.GetSingleByCondition(x => x.chinhanh == chinhanh.chinhanh1);
                        listDaily.Add(daily);
                    }
                    break;
                case "MT":
                    var listChinhanhMT = _chinhanhRepository.GetMulti(x => x.nhom == nhom).ToList();
                    foreach (var chinhanh in listChinhanhMT)
                    {
                        var daily = _dmdailyRepository.GetSingleByCondition(x => x.chinhanh == chinhanh.chinhanh1);
                        listDaily.Add(daily);
                    }
                    break;
                case "MB":
                    var listChinhanhMB = _chinhanhRepository.GetMulti(x => x.nhom == nhom).ToList();
                    foreach (var chinhanh in listChinhanhMB)
                    {
                        var daily = _dmdailyRepository.GetSingleByCondition(x => x.chinhanh == chinhanh.chinhanh1);
                        listDaily.Add(daily);
                    }
                    break;
                default:
                    listDaily = _dmdailyRepository.GetAll().ToList();
                    break;

            }
            //if (chinhanh != "")
            //    listDaily = _dmdailyRepository.GetMulti(x => x.chinhanh == chinhanh && x.trangthai).ToList();
            //else
            //    listDaily = _dmdailyRepository.GetAll().ToList();
            return listDaily;
        }

        public IEnumerable<ThongKeDoanhThuViewModel> ThongKeDoanhThuOB(string khoi)
        {
            var listDanhThu = _thongkeRepository.ThongKeDoanhThuOB(khoi);
            return listDanhThu;
        }

        public IEnumerable<ThongKeKhachViewModel> ThongKeSoKhachOB(string khoi)
        {
            var listDanhThu = _thongkeRepository.ThongKeSoKhachOB(khoi);
            return listDanhThu;
        }
    }
}