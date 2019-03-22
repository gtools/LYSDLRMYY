using System;
using System.Collections.Generic;
using GTSharp.Core;

namespace LYSDLRMYY
{
    /// <summary>
    /// 附件文件下载
    /// </summary>
    public class AttachFile
    {
        #region 参数
        /// <summary>
        /// 参数
        /// 0：URL地址
        /// 1：保存文件夹名称
        /// 2：保存文件名称
        /// 3：是否覆盖
        /// </summary>
        List<object> Params { get; set; }
        #endregion

        /// <summary>
        /// 构造函数
        /// </summary>
        public AttachFile()
        {
            Params = new List<object>();
        }

        /// <summary>
        /// 添加参数
        /// </summary>
        /// <param name="o">参数</param>
        public void AddParam(object o)
        {
            Params.Add(o);
        }

        /// <summary>
        /// 清空参数
        /// </summary>
        public void ClearParam()
        {
            Params.Clear();
        }

        /// <summary>
        /// Http文件下载
        /// 参数
        /// 0：URL地址
        /// 1：保存文件夹名称
        /// 2：保存文件名称
        /// 3：是否覆盖
        /// </summary>
        public void HttpFileDownload()
        {
            //URL地址
            string url = Params[0].ToString();
            //保存文件夹名称
            string dirname = Params[1].ToString();
            //保存文件名称
            string filename = Params[2].ToString();
            //是否覆盖
            bool b = bool.Parse(Params[3].ToString());
            Helper h = new Helper();
            string file = h.PathFile(dirname, filename);
            string dir = h.PathDir(dirname);
            FileHelper.DirCreate(dir);
            //如果存在并且需要覆盖的
            if (b || !FileHelper.FileExists(file))
                HttpHelper.HttpFileDownload(url, file);
        }
    }
}