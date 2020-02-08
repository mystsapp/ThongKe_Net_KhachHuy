using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKe.Data.Models
{
    public class TuyentqChiTiet
    {
        [Key]
        public long stt { get; set; }
        public string chinhanh { get; set; }
        public string tuyentq { get; set; }
        public string sgtcode { get; set; }
        public int vetourid { get; set; }
        public DateTime batdau { get; set; }
        public DateTime ketthuc { get; set; }
        public string dailyxuatve { get; set; }
        public int sk { get; set; }
        public decimal doanhthu { get; set; }
    }
}
