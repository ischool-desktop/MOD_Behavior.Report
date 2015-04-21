using FISCA;
using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Behavior.MeritDemeritConfirmation
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string toolName = "班級獎懲記錄明細(確認表)";

            K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = false;
            K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Click += delegate
            {
                MeritDemeritListForm AttList = new MeritDemeritListForm();
                AttList.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["學務相關報表"][toolName].Enable = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.獎懲確認表_權限);
            };

            Catalog detail = RoleAclSource.Instance["班級"]["報表"];
            detail.Add(new RibbonFeature(Permissions.獎懲確認表, toolName));

        }
    }
}
