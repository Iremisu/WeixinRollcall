using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace weixinRollCall.DAL.Model
{
    public class ClassStudent
    {
        public ClassStudent(string StudentName,String StudentID,string StudentClass)
        {
            this.StudentName = StudentName;
            this.StudentID = StudentID;
            this.StudentClass = StudentClass;
        }
        public string ClassID { get; set; }
        public string StudentName { get; set; }
        public string StudentID { get; set; }
        public string StudentClass { get; set; }
        
    }
}