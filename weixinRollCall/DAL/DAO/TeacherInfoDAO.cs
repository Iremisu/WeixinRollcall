using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using weixinRollCall.DAL.Model;
using MySql.Data.MySqlClient;


namespace weixinRollCall.DAL.DAO
{
    public class TeacherInfoDAO:SqlDAO
    {
        /// <summary>
        /// 根据openid获取教师信息
        /// </summary>
        /// <param name="openid">微信端标识</param>
        /// <returns>教师信息</returns>
        public TeacherInfo GetTeacherInfo(string openid)
        {
            string sql = "SELECT `Tea_ID` , `Tea_XingMing` ,`Tea_MiMa`,`Tea_Email`  FROM `xgteacherinfo` WHERE `openid` ='" + openid+"'";
            MySqlDataReader dr = ExcuteReader(sql);
            return GetTeacherInfo(dr);
        }
        /// <summary>
        /// 根据工号获取教师信息
        /// </summary>
        /// <param name="Tea_ID"></param>
        /// <returns></returns>
        public TeacherInfo GetTeacherbyID (string Tea_ID)
        {
            string sql= "SELECT `Tea_ID` , `Tea_XingMing` ,`Tea_MiMa` ,`Tea_Email` FROM `xgteacherinfo` WHERE `Tea_ID` ='" + Tea_ID + "'";
            MySqlDataReader dr = ExcuteReader(sql);
            return GetTeacherInfo(dr);
        }
        /// <summary>
        /// 根据DataReader取得实例
        /// </summary>
        /// <param name="">DataReader</param>
        /// <returns>实例</returns>
        public TeacherInfo GetTeacherInfo(MySqlDataReader dr)
        {
            if (dr.HasRows)
            {
                dr.Read();
                TeacherInfo teacher = new TeacherInfo(GetString(dr["Tea_ID"]),
                    GetString(dr["Tea_XingMing"]), GetString(dr["Tea_MiMa"]), GetString(dr["Tea_Email"]));
                ConnClose();
                return teacher;
            }
            else
            {
                ConnClose();
                return new TeacherInfo("", "", "", "");
            }
        }
        public void UpdateopenId(TeacherInfo ti)
        {
            string sql = "UPDATE `xgteacherinfo`  SET `openid` ='"+ti.openid+"' WHERE `Tea_ID` ='"+ti.Tea_ID+"'";
            ExexuteNonQuery(sql);
            ConnClose();
        }
    }
}