using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace weixinRollCall.DAL.Model
{
    public class TeacherClass
    {
        public TeacherClass(string ClassID,string ClassName,string TeacherID,string ClassHead, string ClassType,string TotalStuNum,string TimeAndLocation,int LessonNum)
        {
            this.ClassID = ClassID;
            this.ClassName = ClassName;
            this.TeacherID = TeacherID;
            this.ClassHead = ClassHead;
            this.ClassType = ClassType;
            this.TotalStuNum = TotalStuNum;
            this.TimeAndLocation = TimeAndLocation;
            this.LessonNum = LessonNum;
        }
        public string ClassID { get; set; }
        public string ClassName { get; set; }
        public string TeacherID { get; set; }
        public string ClassHead { get; set; }
        public string ClassType { get; set; }
        public string TotalStuNum { get; set; }
        public string TimeAndLocation { get; set; }   
        public int LessonNum { get; set; } 

    }
}