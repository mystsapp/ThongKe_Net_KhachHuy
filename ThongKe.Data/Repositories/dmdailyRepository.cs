using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Infrastructure;
using ThongKe.Data.Models.EF;

namespace ThongKe.Data.Repositories
{
    public interface IdmdailyRepository : IRepository<dmdaily> { }
    public class dmdailyRepository:RepositoryBase<dmdaily>,IdmdailyRepository
    {
        public dmdailyRepository(IDbFactory dbFactory) : base(dbFactory)
        {

        }
    }
}
