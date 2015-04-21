using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Behavior.MeritDemeritConfirmation
{
    class Permissions
    {
        public static string 獎懲確認表 { get { return "Report.Behavior.Class.MeritDemeritListForm.cs"; } }

        /// <summary>
        /// 班級缺曠記錄明細
        /// </summary>
        public static bool 獎懲確認表_權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[獎懲確認表].Executable;
            }
        }
    }
}
