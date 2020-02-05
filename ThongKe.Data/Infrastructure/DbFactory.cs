using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Models;
using ThongKe.Data.Models.EF;

namespace ThongKe.Data.Infrastructure
{
    public class DbFactory : Disposable, IDbFactory
    {
        private TKDbContext dbContext;

        public TKDbContext Init()
        {
            return dbContext ?? (dbContext = new TKDbContext());
        }

        protected override void DisposeCore()
        {
            if (dbContext != null)
                dbContext.Dispose();
        }
    }
}
