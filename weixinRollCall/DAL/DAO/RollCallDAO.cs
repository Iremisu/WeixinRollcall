using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using weixinRollCall.DAL.Model;
using MySql.Data.MySqlClient;

namespace weixinRollCall.DAL.DAO
{
    public class RollCallDAO:SqlDAO
    {
        /// <summary>
        /// 向数据库添加点名信息
        /// </summary>
        /// <param name="rc">点名信息</param>
        public void Edit(List<RollCallInfo> rc)
        {
            foreach (RollCallInfo RC in rc)
            {
                string sql = "INSERT INTO `rollcall`(`StudentID` ,`StudentName` ,`ClassID` ,`Time` ,`Status` ,`TeacherName`,`KKLessonNum`) values('" + RC.StudentID+"','"
                    + RC.StudentName+"','"+ RC.ClassID + "','" + RC.Time + "','"+RC.Status+ "','"+ HttpContext.Current.Session["TeacherName"].ToString()+ "',"+RC.KKLessonNum+")";
                ExexuteNonQuery(sql);
            }
            ConnClose();
        }
        /// <summary>
        /// 获取课程点名所有日期
        /// </summary>
        /// <param name="cid">课程编号</param>
        public List<string> GetDate(string cid)
        {
            string sql = "SELECT DISTINCT(`Time`) FROM `rollcall` WHERE `ClassID` ='" + cid +  "'";
            List<string> date = new List<string>();
            MySqlDataReader dr = ExcuteReader(sql);
            while (dr.Read())
            {
                date.Add(GetString(dr["Time"]));
            }
            ConnClose();
            return date;
        }
        public List<string> GetStatus(List<ClassStudent> cs,string date, string cid)
        {
            string sql;
            List<string> s=new List<string>();
            foreach (ClassStudent CS in cs)
            {
                sql = "SELECT `Status` FROM `rollcall` WHERE `ClassID` ='" + cid + "' and `Time` ='" +date+"' and `StudentName` ='" + CS.StudentName + "'";
                MySqlDataReader dr = ExcuteReader(sql);
                if (dr.Read())
                {
                    s.Add(GetString(dr["Status"]));
                }
                else
                {
                    s.Add("");
                }
                DrClose();
            }
            ConnClose();
            return s;
        }
    }
}