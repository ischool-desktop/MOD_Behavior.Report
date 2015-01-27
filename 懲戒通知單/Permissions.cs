using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.懲戒通知單2013
{
    class Permissions
    {
        public static string 學生懲戒通知單 { get { return "SHSchool.Behavior.Student.DisciplineNotificationForm.2013"; } }
        public static string 班級懲戒通知單 { get { return "SHSchool.Behavior.Class.DisciplineNotificationForm.2013"; } }

        public static bool 學生懲戒通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生懲戒通知單].Executable;
            }
        }

        public static bool 班級懲戒通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級懲戒通知單].Executable;
            }
        }

    }
}
