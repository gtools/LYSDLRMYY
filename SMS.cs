using System;
using System.Linq;
using GTSharp;

namespace LYSDLRMYY
{
    /// <summary>
    /// 短信
    /// </summary>
    public class SMS
    {
        /// <summary>
        /// 电话号码去重
        /// </summary>
        /// <param name="msg">内容</param>
        /// <returns>电话</returns>
        public string TelDistinct(string msg)
        {
            return string.Join(Environment.NewLine, msg.GetSplitByNewLine().Distinct().ToArray());
        }

        /// <summary>
        /// 电话号码长度为11位的
        /// </summary>
        /// <param name="msg">内容</param>
        /// <returns>电话</returns>
        public string TelLength(string msg)
        {
            return string.Join(Environment.NewLine, msg.GetSplitByNewLine(false).Where(t => t.Length == 11).ToArray());
        }

        /// <summary>
        /// 电话号码转换数组
        /// </summary>
        /// <param name="msg">内容</param>
        /// <returns>数组</returns>
        public string[] GetSplitByNewLine(string msg)
        {
            return msg.GetSplitByNewLine();
        }
    }
}