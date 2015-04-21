using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using FISCA.DSAUtil;
using System.IO;

namespace K12.Behavior.AttendanceConfirmation
{
    public class AttendanceConfigData
    {
        /// <summary>
        /// 設定之"NamesPace"名稱
        /// </summary>
        public string ConfigDataString = "班級缺曠明細確認表_Word";
        K12.Data.Configuration.ConfigData cd;

        /// <summary>
        /// 目前系統設定,全部的缺曠別
        /// </summary>
        public List<string> AllAbsence = new List<string>();

        /// <summary>
        /// 範本設定
        /// </summary>
        XmlElement PrintData;
        /// <summary>
        /// 假別設定
        /// </summary>
        XmlElement AttendanceData;

        /// <summary>
        /// 範本模式
        /// </summary>
        public string Setup_Mode = "false";
        /// <summary>
        /// 自定範本內容
        /// </summary>
        public string Setup_Temp;
        public byte[] Temp;
        public string ClassNoData;
        public MemoryStream Setup_template = null;
        public MemoryStream Setup_defaultTemplate = new MemoryStream(Properties.Resources.班級缺曠明細確認表範本);
        public List<string> Setup_AttendanceList = new List<string>();

        /// <summary>
        /// 建構子
        /// </summary>
        public AttendanceConfigData()
        {
            //取得系統設定之節次內容
            GetAbsenceList();

            cd = K12.Data.School.Configuration[ConfigDataString];
            PrintData = cd.GetXml("列印設定", null);

            //<Print>
            //    <PrintMode bool=\"false\"></PrintMode>
            //    <PrintTemp Temp=\"\"></PrintTemp>
            //</Print>
            if (PrintData == null) //如果沒有設定檔
            {
                DSXmlHelper XmlPrint = new DSXmlHelper("Print");

                XmlPrint.AddElement("PrintMode");
                XmlPrint.SetAttribute("PrintMode", "bool", "false");
                XmlPrint.AddElement("PrintTemp");
                XmlPrint.SetAttribute("PrintTemp", "Temp", "");
                XmlPrint.AddElement("PrintClassSet");
                XmlPrint.SetAttribute("PrintClassSet", "Class", "True");

                PrintData = XmlPrint.BaseElement;
            }

            AttendanceData = cd.GetXml("假別設定", null);
            //<AttendanceType>
            //    <AttendanceList List=\"曠課\"></AttendanceList>
            //    <AttendanceList List=\"事假\"></AttendanceList>
            //    <AttendanceList List=\"病假\"></AttendanceList>
            //    <AttendanceList List=\"喪假\"></AttendanceList>
            //    <AttendanceList List=\"公假\"></AttendanceList>
            //    <AttendanceList List=\"婚假\"></AttendanceList>
            //    <AttendanceList List=\"測假\"></AttendanceList>
            //    <AttendanceList List=\"新假\"></AttendanceList>
            //</AttendanceType>

            if (AttendanceData == null) //如果沒有設定檔
            {
                DSXmlHelper XmlAttendance = new DSXmlHelper("AttendanceType");
                foreach (string each in AllAbsence)
                {
                    XmlAttendance.AddElement("AttendanceList");
                    XmlAttendance.SetAttribute("AttendanceList", "List", each);
                }

                AttendanceData = XmlAttendance.BaseElement;

            }

            Reset();

        }

        private void GetAbsenceList()
        {
            foreach (K12.Data.AbsenceMappingInfo each in K12.Data.AbsenceMapping.SelectAll())
            {
                if (!AllAbsence.Contains(each.Name))
                {
                    AllAbsence.Add(each.Name);
                }

            }
        }

        public void Reset()
        {
            XmlElement P = (XmlElement)PrintData.SelectSingleNode("PrintMode");

            if (P != null)
            {
                Setup_Mode = P.GetAttribute("bool");
            }
            else
            {
                Setup_Mode = "false";
            }

            XmlElement T = (XmlElement)PrintData.SelectSingleNode("PrintTemp");

            if (T != null)
            {
                Setup_Temp = T.GetAttribute("Temp");
            }
            else
            {
                Setup_Temp = "";
            }

            XmlElement K = (XmlElement)PrintData.SelectSingleNode("PrintClassSet");

            if (K != null)
            {
                ClassNoData = K.GetAttribute("Class");
            }
            else
            {
                ClassNoData = "True";
            }

            if (Setup_Temp != "")
            {
                Temp = Convert.FromBase64String(Setup_Temp);
                Setup_template = new MemoryStream(Temp);
            }

            Setup_AttendanceList.Clear();
            foreach (XmlNode each in AttendanceData.SelectNodes("AttendanceList"))
            {
                XmlElement list = (XmlElement)each;
                if(!Setup_AttendanceList.Contains(list.GetAttribute("List")))
                {
                    Setup_AttendanceList.Add(list.GetAttribute("List"));
                }
                //list.GetAttribute("List");

                //each.GetAttribute("List");
            }
        }

        /// <summary>
        /// 儲存列印設定
        /// </summary>
        /// <param name="print">傳入列印設定的Xml,並儲存</param>
        public void SavePrint(string mode,string base64,string SetClass)
        {
            Setup_Mode = mode; //模式

            DSXmlHelper XmlPrint = new DSXmlHelper("Print");
            XmlPrint.AddElement("PrintMode");
            XmlPrint.SetAttribute("PrintMode", "bool", Setup_Mode);
            XmlPrint.AddElement("PrintTemp");
            XmlPrint.SetAttribute("PrintTemp", "Temp", base64);
            XmlPrint.AddElement("PrintClassSet");
            XmlPrint.SetAttribute("PrintClassSet", "Class", SetClass);
            PrintData = XmlPrint.BaseElement;
            SaveAll();
        }

        /// <summary>
        /// 儲存假別設定
        /// </summary>
        /// <param name="Attendance">傳入假別設定的Xml,並儲存</param>
        public void SaveAttendance(List<string> Attendance)
        {
            DSXmlHelper XmlAttendance = new DSXmlHelper("AttendanceType");
            foreach (string each in Attendance)
            {
                XmlAttendance.AddElement("AttendanceList");
                XmlAttendance.SetAttribute("AttendanceList", "List", each);
            }

            AttendanceData = XmlAttendance.BaseElement;
            SaveAll();
        }

        public void SaveAll()
        {
            cd["列印設定"] = PrintData.OuterXml;
            cd["假別設定"] = AttendanceData.OuterXml;
            cd.Save();
            Reset();
        }
    }
}
