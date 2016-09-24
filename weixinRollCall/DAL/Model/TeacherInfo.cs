using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace weixinRollCall.DAL.Model
{
    public class TeacherInfo
    {
        public TeacherInfo(string TEA_ID,string TEA_XingMing,string Tea_MiMa,string Tea_Email)
        {
            this.Tea_ID = TEA_ID;
            this.Tea_XingMing = TEA_XingMing;
            this.Tea_MiMa = Tea_MiMa;
            this.Tea_Email = Tea_Email;
        }
        public string Tea_ID { get; set; }
        public string openid { get; set; }
        public string Tea_XingMing { get; set; }  
        public string Tea_MiMa { get; set; } 
        public string Tea_Email { get; set; }

    }
}