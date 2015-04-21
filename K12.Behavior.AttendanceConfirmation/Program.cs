using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using K12.Data;
using FISCA.Permission;

namespace K12.Behavior.AttendanceConfirmation
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string toolName = "班級缺曠記錄明細(確認表)";

            K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = false;
            K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Click += delegate
            {
                AttendanceListForm AttList = new AttendanceListForm();
                AttList.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.班級缺曠記錄明細權限);
            };

            Catalog detail = RoleAclSource.Instance["班級"]["報表"];
            detail.Add(new RibbonFeature(Permissions.班級缺曠記錄明細, toolName));

        }
    }
}
