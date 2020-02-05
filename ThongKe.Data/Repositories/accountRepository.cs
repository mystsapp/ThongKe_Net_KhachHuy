using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Infrastructure;
using ThongKe.Data.Models.EF;

namespace ThongKe.Data.Repositories
{
    public interface IaccountRepository : IRepository<account>
    {
    }
    public class accountRepository : RepositoryBase<account>, IaccountRepository
    {
        public accountRepository(IDbFactory dbFactory) : base(dbFactory)
        {
        }

    }
}
