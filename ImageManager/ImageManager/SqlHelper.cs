 using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImageManager
{
    class SqlHelper
    {

        private static string conStr = ConfigurationManager.AppSettings["DBConstr"];

        //插入,删除数据时 调用这个方法
        public static int ExecuteNonQuery(string sql,params SqlParameter [] parameters)
        {
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    return cmd.ExecuteNonQuery();
                }
            }

        }
        public static object ExecuteScalar(string sql , params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    return cmd.ExecuteScalar();
                }
            }
        }
        //查询时调用这个方法 ，返回一个dataset
        public static DataSet ExecuteDataSet(string sql,params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(conStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                   // using (SqlDataReader reader = cmd.ExecuteReader())
                     SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataSet dataset = new DataSet();
                        adapter.Fill(dataset);
                       // DataTable DT = new DataTable();
                       // DT.Load(reader);
                        return dataset;

                    
                }
            }
        }
    }
}
