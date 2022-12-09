using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient; //ket noi
using System.Data; //thu vien dataset

namespace DoAnLapTrinh
{
    class KetNoi
    {
        public string con_str = "";
        public SqlConnection conn = null;
        public KetNoi()
        {
            con_str = @"Data Source=ADMIN\SQLEXPRESS;Initial Catalog=DoAnLapTrinh;Integrated Security=True";
            conn = new SqlConnection(con_str);
        }
        public bool thucthi(string sql)
        {
            try
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public DataSet laydulieu(string sql, string table_name)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                da.Fill(ds, table_name);
            }
            catch
            {
            }
            return ds;
        }
    }
}

