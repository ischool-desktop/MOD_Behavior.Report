using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.銷過通知單2015
{
    class Permissions
    {
        public static string 學生銷過通知單 { get { return "Behavior.Student.ClearDemeritReport.2013"; } }
        public static string 班級銷過通知單 { get { return "Behavior.Class.ClearDemeritReport.2013"; } }

        public static bool 學生銷過通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生銷過通知單].Executable;
            }
        }

        public static bool 班級銷過通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級銷過通知單].Executable;
            }
        }

    }
}
