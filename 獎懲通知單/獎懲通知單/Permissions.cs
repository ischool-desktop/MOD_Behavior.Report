using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 獎懲通知單
{
    class Permissions
    {
        public static string 學生獎懲通知單 { get { return "SHSchool.Behavior.Student.MeritDisciplineNotificationForm.2013"; } }
        public static string 班級獎懲通知單 { get { return "SHSchool.Behavior.Class.MeritDisciplineNotificationForm.2013"; } }

        public static bool 學生獎懲通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生獎懲通知單].Executable;
            }
        }

        public static bool 班級獎懲通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級獎懲通知單].Executable;
            }
        }

    }

}
