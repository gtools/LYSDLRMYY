using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace LYSDLRMYY
{
    /// <summary>
    /// 全院未提交病历
    /// </summary>
    public class QYWTJBL
    {
        #region 参数
        /// <summary>
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        List<object> Params { get; set; }
        #endregion

        /// <summary>
        /// 构造函数
        /// </summary>
        public QYWTJBL()
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
        /// 导出模板：每周1全院未交病历.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void WeekReport1()
        {
            //数据
            DataTable data = (DataTable)Params[0];
            //写入数据开始行
            int beginrow = int.Parse(Params[1].ToString());
            //增加天数
            int addday = int.Parse(Params[2].ToString());
            //模板文件夹
            string template_dir = Params[3].ToString();
            //模板文件
            string template_file = Params[4].ToString();
            //保存文件夹
            string temp_dir = Params[5].ToString();
            //保存文件
            string temp_file = Params[6].ToString();
            //帮助
            Helper h = new Helper();
            //模版路径
            string template_dir_path = h.PathDir(template_dir);
            string template_file_path = h.PathFile(template_dir, template_file);
            //保存路径
            string temp_dir_path = h.PathDir(temp_dir);
            string temp_file_path = h.PathFile(temp_dir, temp_file);
            GTSharp.Core.FileHelper.DirCreate(template_dir_path, false);
            GTSharp.Core.FileHelper.DirCreate(temp_dir_path, false);
            //清空保存目录
            h.DirClear(temp_dir);
            //写入数据结束行
            int endrow = beginrow + data.Rows.Count - 1;
            //写入数据结束行
            int column = data.Columns.Count;
            //临时数据
            string _temp = string.Empty;
            //数据分组
            DataView dataview = data.DefaultView;
            //注：其中ToTable（）的第一个参数为是否DISTINCT
            DataTable distinct = dataview.ToTable(true, "科室");
            foreach (DataRow item in distinct.Rows)
            {
                DataTable dttemp = h.DataTableSelect(data, "科室='" + item[0].ToString() + "'", "住院号");
                //没数据返回
                if (dttemp.Rows.Count <= 0)
                    continue;
                //导入模版
                GTSharp.Core.ExcelHelper exl = new GTSharp.Core.ExcelHelper(template_file_path);
                //设置单元格日期
                h.SetReplace_DATE(addday, exl);
                //设置单元格日期
                h.SetReplace_NUM(dttemp.Rows.Count, exl.GetCell(1, 2), exl);
                //导出数据到Excel
                exl.DataTableToExcel(dttemp, beginrow);
                //写入数据结束行
                endrow = beginrow + dttemp.Rows.Count - 1;
                //添加边框
                exl.StyleLine(exl.GetRange(beginrow, 1, endrow, column));
                //保存
                exl.SaveAsFile(temp_dir_path + item[0].ToString());
                Thread.Sleep(50);
            }
        }
    }
}