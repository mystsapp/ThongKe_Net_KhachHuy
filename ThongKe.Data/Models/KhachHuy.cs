using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThongKe.Data.Models
{
    public class KhachHuy
    {
        public long stt { get; set; }
        public string tenkhach { get; set; }
        public string sgtcode { get; set; }
        public int vetourid { get; set; }
        public string tuyentq { get; set; }
        public DateTime batdau { get; set; }
        public DateTime ketthuc { get; set; }
        public decimal giatour { get; set; }
        public string nguoihuyve { get; set; }
        public string dailyhuyve { get; set; }
        public string chinhanh { get; set; }
        public DateTime ngayhuyve { get; set; }
    }
}
