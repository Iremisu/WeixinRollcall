using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using weixinRollCall.DAL.Model;
using MySql.Data.MySqlClient;
using System.Web.UI;
using weixinRollCall.Common;


namespace weixinRollCall.DAL.DAO
{
    public class TeacherClassDAO : SqlDAO
    {
        /// <summary>
        /// 根据教师ID取得课程信息
        /// </summary>
        /// <param name="TeacherName">教师ID</param>
        /// <returns>课程ID</returns>
        public List<TeacherClass> GetClass(string TeacherID)
        {
            string sql = "SELECT * FROM `teacherclass` WHERE `TeacherID` ='"+TeacherID+ "' and `ClassID` like '"+Util.GetTerm()+"%'";
            List<TeacherClass> tc = new List<TeacherClass>();            
            MySqlDataReader dr = ExcuteReader(sql);
            while(dr.Read())
            {
                tc.Add(GetClass(dr));
            }
            ConnClose();
            return tc;
        }
        internal static TeacherClass GetClass(MySqlDataReader dr)
        {            
            TeacherClass tc = new TeacherClass(GetString(dr["ClassID"]), GetString(dr["ClassName"]), 
                GetString(dr["TeacherID"]), GetString(dr["ClassHead"]), GetString(dr["ClassType"]), 
                GetString(dr["TotalStuNum"]), GetString(dr["TimeAndLocation"]),GetInt32(dr["LessonNum"]));
            return tc;
        }
    }
}