using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using weixinRollCall.DAL.Model;
using weixinRollCall.DAL.DAO;
using weixinRollCall.Common;

namespace weixinRollCall
{
    public partial class RollCall : System.Web.UI.Page
    {
        string ClassID = "";
        protected string ClassName = "";
        List<RollCallInfo> rc = new List<RollCallInfo>();
        List<ClassStudent> cs = new List<ClassStudent>();
        
        int LessonNum;
        string excelpath = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            ClassName = Request.QueryString["ClassName"];
            ClassID = Request.QueryString["ClassID"];
            LessonNum= Int32.Parse(Request.QueryString["LessonNum"]);
            excelpath = "C:\\Excel\\" + ClassID + ".xls";
            cs = new ClassStudentDAO().GetStudent(ClassID);
            this.Repeater1.DataSource = cs;
            this.Repeater1.DataBind();
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            int arrival=0, absence=0, events=0;
            DateTime dt = DateTime.Now;
            string TeacherID = Session["TeacherID"].ToString();
            foreach (ClassStudent CS in cs)
            {
                string status= Request.Form[CS.StudentID];
                
                if (!string.IsNullOrEmpty(status))
                {
                    if (status == "未到")
                    {
                        rc.Add(new RollCallInfo(CS.StudentID, CS.StudentName, ClassID, dt.ToString("yyyyMMdd HH:mm:ss"), status, LessonNum));
                        absence++;
                    }
                    else if(status=="已到")
                    {
                        rc.Add(new RollCallInfo(CS.StudentID, CS.StudentName, ClassID, dt.ToString("yyyyMMdd HH:mm:ss"), status, 0));
                        arrival++;
                    }
                    else
                    {
                        rc.Add(new RollCallInfo(CS.StudentID, CS.StudentName, ClassID, dt.ToString("yyyyMMdd HH:mm:ss"), status, 0));
                        events++;
                    }
                }
            }
                new RollCallDAO().Edit(rc);
            Page.ClientScript.RegisterStartupScript(Page.GetType(), 
                "message", "<script type='text/javascript'>alert('点名成功！');location.href='selectclass.aspx';</script>");            
        }
      }
}