using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MySql.Data.MySqlClient;
using weixinRollCall.DAL.Model;

namespace weixinRollCall.DAL.DAO
{
    public class ClassStudentDAO:SqlDAO
    {
        /// <summary>
        /// 根据课程编号获取名单
        /// </summary>
        /// <param name="ClassID">课程编号</param>
        /// <returns>名单</returns>
        public List<ClassStudent> GetStudent(string ClassID)
        {
            string sql = "SELECT * FROM `classstudent` WHERE `ClassID` = " + ClassID + " order by `StudentID` ASC";
            MySqlDataReader dr = ExcuteReader(sql);
            List<ClassStudent> cs = new List<ClassStudent>();
            while(dr.Read())
            {
                cs.Add(GetStudent(dr));
            }
            ConnClose();
            return cs;
        }
        internal static ClassStudent GetStudent(MySqlDataReader dr)
        {
            ClassStudent cs = new ClassStudent(GetString(dr["StudentName"]), 
                GetString(dr["StudentID"]), GetString(dr["StudentClass"]));
            return cs;
        }
    }
}