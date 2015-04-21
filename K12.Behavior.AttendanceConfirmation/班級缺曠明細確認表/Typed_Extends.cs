using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using K12.Data.Configuration;
using K12.Data;

namespace K12.Behavior.AttendanceConfirmation
{
    /// <summary>
    /// 延伸 ConfigData 有關型別處理的功能。
    /// </summary>
    public static class Typed_Extends
    {
        /// <summary>
        /// 取得 Boolean 資料。如果資料不存在會回傳「defaultValue」，如果資料存在會轉型成 Boolean。轉型失敗會產生 Exception。
        /// </summary>
        public static bool GetBoolean(this ConfigData config, string name, bool defaultValue)
        {
            if (config.Contains(name))
            {
                bool result;
                if (bool.TryParse(config[name], out result))
                    return result;
                else
                    throw new ArgumentException(string.Format("「{0}」資料無法轉型成 Boolean 型別。", name));
            }
            else
                return defaultValue;
        }

        /// <summary>
        /// 設定 Boolean 資料。
        /// </summary>
        public static void SetBoolean(this ConfigData config, string name, bool value)
        {
            config[name] = value.ToString();
        }

        /// <summary>
        /// 取得 Integer 資料。如果資料不存在會回傳「defaultValue」，如果資料存在會轉型成 Integer。轉型失敗會產生 Exception。
        /// </summary>
        public static int GetInteger(this ConfigData config, string name, int defaultValue)
        {
            if (config.Contains(name))
            {
                int result;
                if (int.TryParse(config[name], out result))
                    return result;
                else
                    throw new ArgumentException(string.Format("「{0}」資料無法轉型成 Int 型別。", name));
            }
            else
                return defaultValue;
        }

        /// <summary>
        /// 設定 Integer 資料。
        /// </summary>
        public static void SetInteger(this ConfigData config, string name, int value)
        {
            config[name] = value.ToString();
        }

        /// <summary>
        /// 取得 Xml 資料。如果資料不存在會回傳「defaultValue」，如果資料存在會轉型成 XmlElement。轉型失敗會產生 Exception。
        /// </summary>
        public static XmlElement GetXml(this ConfigData config, string name, XmlElement defaultValue)
        {
            if (config.Contains(name))
                return XmlHelper.LoadXml(config[name]);
            else
                return defaultValue;
        }

        /// <summary>
        /// 設定 Xml 資料。
        /// </summary>
        public static void SetXml(this ConfigData config, string name, XmlElement value)
        {
            if (value == null)
                throw new ArgumentException("參數不可以是 Null。","value");

            config[name] = value.OuterXml;
        }
    }
}
