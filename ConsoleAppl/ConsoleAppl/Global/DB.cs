using System.Configuration;
using System.Data;
using System.Data.SqlClient;
namespace ConsoleAppl.Global
{
    /// <summary>
    /// 数据库链接
    /// </summary>
    public class DB
    {
        
        public static DataSet getDataSet(string sql) { 
            string constr = ConfigurationManager.ConnectionStrings["DBSetting"].ConnectionString.ToString();
            if(string.IsNullOrEmpty(sql)) return new DataSet();
            SqlConnection con = new SqlConnection(constr);

            try
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(sql, con);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch
            {
                return new DataSet();
            }
            finally 
            {
                con.Close();
            }
        }
    }
}
