using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using FISCA.Presentation.Controls;

namespace K12.懲戒通知單2015
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            string URL學生懲戒通知單 = "ischool/高中系統/共用/學務/學生/報表/懲戒通知單_2013";
            string URL班級懲戒通知單 = "ischool/高中系統/共用/學務/班級/報表/懲戒通知單_2013";

            string toolName = "2015懲戒通知單(測試版)";

            FISCA.Features.Register(URL學生懲戒通知單, arg =>
               {
                   if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                   {
                       new Report("student").Print();
                   }
                   else
                   {
                       MsgBox.Show("產生學生報表,請選擇學生!!");
                   }
               });

            RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            StudentReports["報表"]["學務相關報表"][toolName].Enable = false;
            StudentReports["報表"]["學務相關報表"][toolName].Click += delegate
            {
                Features.Invoke(URL學生懲戒通知單);
            };

            FISCA.Features.Register(URL班級懲戒通知單, arg =>
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    new Report("class").Print();
                }
                else
                {
                    MsgBox.Show("產生班級報表,請選擇班級!!");
                }
            });

            RibbonBarItem ClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            ClassReports["報表"]["學務相關報表"][toolName].Enable = false;
            ClassReports["報表"]["學務相關報表"][toolName].Click += delegate
            {
                Features.Invoke(URL班級懲戒通知單);
            };

            //學生選擇
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                {
                    StudentReports["報表"]["學務相關報表"][toolName].Enable = false;
                }
                else
                {
                    StudentReports["報表"]["學務相關報表"][toolName].Enable = Permissions.學生懲戒通知單權限;
                }
            }; 

            //班級選擇
            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count <= 0)
                {
                    ClassReports["報表"]["學務相關報表"][toolName].Enable = false;
                }
                else
                {
                    ClassReports["報表"]["學務相關報表"][toolName].Enable = Permissions.班級懲戒通知單權限;
                }
            }; 

            Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.學生懲戒通知單, toolName));

            ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.班級懲戒通知單, toolName));
        }
    }
}
