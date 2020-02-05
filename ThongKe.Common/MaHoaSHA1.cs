using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace ThongKe.Common
{
    public class MaHoaSHA1
    {
        public string EncodeSHA1(string pass)
        {
            SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider();
            byte[] bs = System.Text.Encoding.UTF8.GetBytes(pass);
            bs = sha1.ComputeHash(bs);
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            foreach (byte b in bs)
            {
                s.Append(b.ToString("x1").ToUpper());
            }
            pass = s.ToString();
            return pass;
        }
    }
    public static class EntityToTable
    {
        public static DataTable ToDataTable<T>(this IEnumerable<T> entityList) where T : class
        {
            try
            {
                var properties = typeof(T).GetProperties();
                var table = new DataTable();

                foreach (var property in properties)
                {
                    var type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                    table.Columns.Add(property.Name, type);
                }
                foreach (var entity in entityList)
                {
                    table.Rows.Add(properties.Select(p => p.GetValue(entity, null)).ToArray());
                }
                return table;
            }
            catch
            {
                return null;
            }
        }



        //internal static DataTable ToDataTable(int result)
        //{
        //    throw new NotImplementedException();
        //}
    }

}
