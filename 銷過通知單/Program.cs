using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using FISCA.Presentation.Controls;

namespace K12.銷過通知單2015
{
    //本功能
    //同時可提供給國高中使用
    //但是需注意
    //權限控管及功能位置需調整開放

    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string URL學生銷過通知單 = "ischool/高中系統/共用/學務/學生/報表/銷過通知單_2013";
            string URL班級銷過通知單 = "ischool/高中系統/共用/學務/班級/報表/銷過通知單_2013";

            string toolName = "銷過通知單(測試版)";

            FISCA.Features.Register(URL學生銷過通知單, arg =>
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    new K12.銷過通知單2015.Report("student").Print();
                }
                else
                {
                    MsgBox.Show("產生學生報表,請選擇學生!!");
                }
            });

            FISCA.Features.Register(URL班級銷過通知單, arg =>
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    new K12.銷過通知單2015.Report("class").Print();
                }
                else
                {
                    MsgBox.Show("產生班級報表,請選擇班級!!");
                }
            });

            #region 高中

            //學生
            RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            StudentReports["報表"]["學務相關報表"][toolName].Enable = false;
            StudentReports["報表"]["學務相關報表"][toolName].Click += delegate
            {
                Features.Invoke(URL學生銷過通知單);
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                {
                    StudentReports["報表"]["學務相關報表"][toolName].Enable = false;
                }
                else
                {
                    StudentReports["報表"]["學務相關報表"][toolName].Enable = Permissions.學生銷過通知單權限;
                }
            };

            //班級
            RibbonBarItem ClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            ClassReports["報表"]["學務相關報表"][toolName].Enable = false;
            ClassReports["報表"]["學務相關報表"][toolName].Click += delegate
            {
                Features.Invoke(URL班級銷過通知單);
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count <= 0)
                {
                    ClassReports["報表"]["學務相關報表"][toolName].Enable = false;
                }
                else
                {
                    ClassReports["報表"]["學務相關報表"][toolName].Enable = Permissions.班級銷過通知單權限;
                }
            };


            //高中權控
            Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.學生銷過通知單, toolName));
            ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.班級銷過通知單, toolName));

            #endregion

            #region 國中註解

            //RibbonBarItem JHClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            //JHClassReports["報表"]["學務相關報表"]["銷過通知單"].Enable = Permissions.班級銷過通知單權限;
            //JHClassReports["報表"]["學務相關報表"]["銷過通知單"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
            //    {
            //        new K12.銷過通知單.Report("class").Print();
            //    }
            //    else
            //    {
            //        MsgBox.Show("產生班級報表,請選擇班級!!");
            //    }
            //};

            //RibbonBarItem JHStudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            //JHStudentReports["報表"]["學務相關報表"]["銷過通知單"].Enable = Permissions.學生銷過通知單權限;
            //JHStudentReports["報表"]["學務相關報表"]["銷過通知單"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
            //    {
            //        new K12.銷過通知單.Report("student").Print();
            //    }
            //    else
            //    {
            //        MsgBox.Show("產生學生報表,請選擇學生!!");
            //    }
            //};

            //Framework.Security.Catalog ribbon2 = Framework.Security.RoleAclSource.Instance["學生"]["報表"];
            //ribbon2.Add(new Framework.Security.RibbonFeature(Permissions.學生銷過通知單, "銷過通知單"));
            //ribbon2 = Framework.Security.RoleAclSource.Instance["班級"]["報表"];
            //ribbon2.Add(new Framework.Security.RibbonFeature(Permissions.班級銷過通知單, "銷過通知單"));

            #endregion
        }
    }
}
