using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;

namespace weixinRollCall.Common
{
    public class Util
    {
        /// <summary>
        /// 获得32位的MD5加密
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string GetMD5_32(string input)
        {
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] data = md5.ComputeHash(System.Text.Encoding.Default.GetBytes(input));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sb.Append(data[i].ToString("x2"));
            }
            return sb.ToString();
        }
        /// <summary>
        /// /获取学期
        /// </summary>
        /// <returns></returns>
        public static string GetTerm()
        {
            int Y = DateTime.Now.Year;
            int M = DateTime.Now.Month;
            string m = "";
            int y = Y - 2000;
            m=(M < 9 && M > 2 )? "02" : "01";
            if (m == "01")
                return y.ToString() + (y + 1).ToString() + m;
            else
                return (y - 1).ToString() + y.ToString() + m;
        }
    }
}