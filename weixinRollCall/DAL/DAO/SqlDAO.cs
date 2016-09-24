using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Configuration;

namespace weixinRollCall.DAL.DAO
{
    public class SqlDAO
    {
        static string connStr = ConfigurationManager.ConnectionStrings["SQLConnString"].ToString();
               
        MySqlConnection conn=new MySqlConnection(connStr);
        MySqlDataReader MDr;

        protected MySqlDataReader ExcuteReader(string sql)
        {            
            if (conn.State.ToString() != "Open")
            {
                conn.Open();
            }
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            MDr = cmd.ExecuteReader();                  
            //conn.Close();
            return MDr;
        }
        protected int ExexuteNonQuery(string sql)
        {
            if (conn.State.ToString() != "Open")
            {
                conn.Open();
            }
            MySqlCommand cmd = new MySqlCommand(sql, conn); 
            int res= cmd.ExecuteNonQuery(); ;
            //conn.Close();
            return res;
        }
        /// <summary>
        /// 获取string数据
        /// </summary>
        /// <param name="data">数据</param>
        /// <returns></returns>
        protected static string GetString(object data)
        {
            if (data == DBNull.Value)
                return "";
            else
                return data.ToString();
        }
        /// <summary>
        /// 获取整型数据
        /// </summary>
        /// <param name="data">数据</param>
        protected static int GetInt32(object data)
        {
            if (data == DBNull.Value)
                return 0;
            else
                return Convert.ToInt32(data);
        }
        protected void DrClose()
        {
            MDr.Close();
        }
        protected void ConnClose()
        {
            conn.Close();
        }
    }
}