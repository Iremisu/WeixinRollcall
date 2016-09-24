using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using weixinRollCall.DAL.Model;
using weixinRollCall.DAL.DAO;
using weixinRollCall.Common;

namespace weixinRollCall
{
    public partial class SelectClass : System.Web.UI.Page
    {
        List<TeacherClass> tc = new List<TeacherClass>();
        protected string TeacherName;
        protected string ClassID;
        protected string ClassName;
        string excelpath;
        List<ClassStudent> cs = new List<ClassStudent>();
        protected void Page_Load(object sender, EventArgs e)
        {
            //根据session里的teacherID获取teacherclass 类
            tc = new TeacherClassDAO().GetClass(Session["TeacherID"].ToString());            
            TeacherName = Session["TeacherName"].ToString();
            this.Repeater1.DataSource = tc;
            Repeater1.DataBind();
        }
        protected void Download(object sender, CommandEventArgs e)
        {
            if (string.IsNullOrEmpty(Session["TeacherEmail"].ToString()))
            {
                Page.ClientScript.RegisterStartupScript(Page.GetType(), "message",
                "<script type='text/javascript'>alert('对不起，您的邮箱没有登记，请将教师姓名、所在学院、邮箱发至cyq @zjut.edu.cn');location.replace(location.href);</script>");
            }
            else
            {
                ClassID = e.CommandArgument.ToString();
                ClassName = e.CommandName.ToString();
                excelpath = "C:\\Excel\\" + ClassID + ".xls";
                cs = new ClassStudentDAO().GetStudent(ClassID);
                Excel ex = new Excel(cs, excelpath, ClassName);//新建模板
                ex.editexcel(ClassID);
                Mail.send(Session["TeacherEmail"].ToString(), "[信息学院点名系统]" + ClassName + "点名情况汇总", excelpath);//Session["TeacherEmail"].ToString();
                Page.ClientScript.RegisterStartupScript(Page.GetType(), "message",
                "<script type='text/javascript'>alert('下载成功！本课程的学生出勤记录已发至"+ Session["TeacherEmail"].ToString()+"邮箱');location.replace(location.href);</script>");
            }
       }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Response.Redirect("register.aspx?openid="+Session["openid"].ToString());
        }
    }
}