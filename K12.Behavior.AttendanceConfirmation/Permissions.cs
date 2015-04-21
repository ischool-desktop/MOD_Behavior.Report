using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.AttendanceConfirmation
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {

        public static string 班級缺曠記錄明細 { get { return "Report.Behavior.Class.Report0020"; } }

        /// <summary>
        /// 班級缺曠記錄明細
        /// </summary>
        public static bool 班級缺曠記錄明細權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級缺曠記錄明細].Executable;
            }
        }
    }
}
