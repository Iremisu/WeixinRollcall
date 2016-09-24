using System;
using System.Security.Cryptography;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using weixinRollCall.DAL.DAO;
using weixinRollCall.Common;
using weixinRollCall.DAL.Model;

namespace weixinRollCall
{
    public partial class Register : System.Web.UI.Page
    {        
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string openid= Request.QueryString["openid"];
                Session["openid"] = openid;
                if (string.IsNullOrEmpty(openid))
                {
                    Response.Redirect("error.html");
                }
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {

            //判断工号密码 是否正确 调用DAL 
            //如果正确 将openid  传入数据库 在session 添加USER
            //Response.Redirect("/SelectClass.aspx");
            //如果错误 提示
            string uid = Request.Form["inputID"];
            string pwd = Request.Form["inputPassword"];
            TeacherInfo ti= new TeacherInfoDAO().GetTeacherbyID(uid);
            string code = ti.Tea_MiMa;
            string md5pwd = Util.GetMD5_32(pwd);
            if (md5pwd == code)//密码正确
            {
                ti.openid = Session["openid"].ToString();
                Session["teacherID"] = ti.Tea_ID;
                Session["TeacherName"] = ti.Tea_XingMing;
                Session["TeacherEmail"] = ti.Tea_Email;
                new TeacherInfoDAO().UpdateopenId(ti);
                Response.Redirect("SelectClass.aspx");
            }
            else//密码错误
            {
                Page.ClientScript.RegisterStartupScript(Page.GetType(), "message", 
                    "<script type='text/javascript'>alert('登录信息错误，请重试！');location.replace(location.href);</script>");
            }
        }
    }
}