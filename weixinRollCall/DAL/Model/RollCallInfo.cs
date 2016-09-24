using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace weixinRollCall.DAL.Model
{
    public class RollCallInfo
    {
        public RollCallInfo(string StudentID, string StudentName, string ClassID, string Time, string Status, int KKLessonNum)
        {
            this.StudentID = StudentID;
            this.StudentName = StudentName;
            this.ClassID = ClassID;
            this.Time = Time;
            this.Status = Status;
            this.KKLessonNum = KKLessonNum;
        }
        public string StudentID { get; set; }
        public string StudentName { get; set; }
        public string ClassID { get; set; }
        public string Time { get; set; }
        public string Status { get; set; }
        public string TeacherID{ get; set; }
        public int KKLessonNum { get; set; }
    }
}