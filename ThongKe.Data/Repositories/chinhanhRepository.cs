using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ThongKe.Data.Infrastructure;
using ThongKe.Data.Models.EF;

namespace ThongKe.Data.Repositories
{
    public interface IchinhanhRepository : IRepository<chinhanh>
    {

    }
    public class chinhanhRepository : RepositoryBase<chinhanh>, IchinhanhRepository
    {
        public chinhanhRepository(IDbFactory dbFactory) : base(dbFactory)
        {

        }
    }
}
