using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Models;
using ThongKe.Data.Models.EF;

namespace ThongKe.Data.Infrastructure
{
    public interface IDbFactory : IDisposable
    {
        TKDbContext Init();
    }
}
