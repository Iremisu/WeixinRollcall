using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using weixinRollCall.DAL.DAO;
using weixinRollCall.DAL.Model;


namespace weixinoepenid
{    
    public partial class WebForm1 : System.Web.UI.Page
    {
        public string code = "";
        string appid = "wx26a116becd3f38c3";
        string secret = "456a92d3e51a3e8d81663f18af38619b";
        string openid = "";
        public void codegetopenid(string code)
        {
            
            string url = string.Format("https://api.weixin.qq.com/sns/oauth2/access_token?appid={0}&secret={1}&code={2}&grant_type=authorization_code",
                appid, secret, code);
            string respText="";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();       
                using (Stream resStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(resStream,Encoding.Default);
                respText = reader.ReadToEnd();
                resStream.Close();
            }

            JObject j=(JObject)JsonConvert.DeserializeObject(respText);
            if (j["openid"] != null)
            {
                openid = j["openid"].ToString();
            }            
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            code = Request.QueryString["code"];
            if (!string.IsNullOrEmpty(code))
            {
                codegetopenid(code);
                TeacherInfo ti = new TeacherInfoDAO().GetTeacherInfo(openid);
                if (ti.Tea_ID == "")
                {
                    Response.Redirect("Register.aspx?openid="+openid);
                }
                else
                {
                    Session["TeacherID"] = ti.Tea_ID;
                    Session["TeacherName"] = ti.Tea_XingMing;
                    Session["TeacherEmail"] = ti.Tea_Email;
                    Session["openid"] = openid;
                    Response.Redirect("SelectClass.aspx");
                }
            }
        }
    }
}