using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using GTSharp;
using GTSharp.Core;

namespace LYSDLRMYY
{
    /// <summary>
    /// 每月院长查询报表
    /// </summary>
    public class MYYZCXBB
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
        public MYYZCXBB()
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
        /// 导出模板：每月1住院主要业务数据同期比表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：开始日期
        /// 3：结束日期
        /// 4：模版文件夹名称
        /// 5：模版文件名称
        /// 6：保存文件夹名称
        /// 7：保存文件名称
        /// </summary>
        public void MonthReport1()
        {
            //数据
            DataTable data = (DataTable)Params[0];
            //写入数据开始行
            int beginrow = int.Parse(Params[1].ToString());
            //开始时间
            DateTime beginday = (DateTime)Params[2];
            DateTime endday = (DateTime)Params[3];
            //模板文件夹
            string template_dir = Params[4].ToString();
            //模板文件
            string template_file = Params[5].ToString();
            //保存文件夹
            string temp_dir = Params[6].ToString();
            //保存文件
            string temp_file = Params[7].ToString();
            //模版路径
            string template_dir_path = Helper.DirPath(template_dir);
            string template_file_path = template_dir_path + template_file;
            //保存路径
            string temp_dir_path = Helper.DirPath(temp_dir);
            string temp_file_path = template_dir_path + temp_file;
            FileHelper.DirCreate(template_dir_path, false);
            FileHelper.DirCreate(temp_dir_path, false);
            //写入数据结束行
            int endrow = beginrow + data.Rows.Count - 1;
            //写入数据结束行
            int column = data.Columns.Count;
            //临时数据
            string _temp = string.Empty;
            //导入模版
            ExcelHelper exl = new ExcelHelper(template_file_path);
            //设置单元格日期
            string strtem1 = string.Format("{0}-{1}月", beginday.ToString("yyyy年M"), endday.Month.ToString());
            string strtem2 = string.Format("{0}-{1}月", beginday.AddYears(-1).ToString("yyyy年M"), endday.Month.ToString());
            string strtem3 = string.Format("{0}年\r\n{1}-{2}月", beginday.Year.ToString(), beginday.Month.ToString(), endday.Month.ToString());
            string strtem4 = string.Format("{0}年\r\n{1}-{2}月", beginday.AddYears(-1).Year.ToString(), beginday.Month.ToString(), endday.Month.ToString());
            //导出数据到Excel
            exl.GetFirst().SetReplace("[DATE1]", strtem1 + "和" + strtem2);
            exl.GetCell(2, 2).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 3).SetReplace("[DATE3]", strtem3);
            exl.GetCell(2, 6).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 7).SetReplace("[DATE3]", strtem3);
            exl.GetCell(2, 10).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 11).SetReplace("[DATE3]", strtem3);
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.GetRange(beginrow, 1, endrow, column).StyleLine();
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //打开
            exl.OpenExcel(temp_file_path);
        }

        /// <summary>
        /// 导出模板：每月2医技科室收入数据同期比表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：开始日期
        /// 3：结束日期
        /// 4：模版文件夹名称
        /// 5：模版文件名称
        /// 6：保存文件夹名称
        /// 7：保存文件名称
        /// </summary>
        public void MonthReport2()
        {
            //数据
            DataTable data = (DataTable)Params[0];
            //写入数据开始行
            int beginrow = int.Parse(Params[1].ToString());
            //开始时间
            DateTime beginday = (DateTime)Params[2];
            DateTime endday = (DateTime)Params[3];
            //模板文件夹
            string template_dir = Params[4].ToString();
            //模板文件
            string template_file = Params[5].ToString();
            //保存文件夹
            string temp_dir = Params[6].ToString();
            //保存文件
            string temp_file = Params[7].ToString();
            //模版路径
            string template_dir_path = Helper.DirPath(template_dir);
            string template_file_path = template_dir_path + template_file;
            //保存路径
            string temp_dir_path = Helper.DirPath(temp_dir);
            string temp_file_path = template_dir_path + temp_file;
            FileHelper.DirCreate(template_dir_path, false);
            FileHelper.DirCreate(temp_dir_path, false);
            //写入数据结束行
            int endrow = beginrow + data.Rows.Count - 1;
            //写入数据结束行
            int column = data.Columns.Count;
            //临时数据
            string _temp = string.Empty;
            //导入模版
            ExcelHelper exl = new ExcelHelper(template_file_path);
            //设置单元格日期
            string strtem1 = string.Format("{0}-{1}月", beginday.ToString("yyyy年M"), endday.Month.ToString());
            string strtem2 = string.Format("{0}-{1}月", beginday.AddYears(-1).ToString("yyyy年M"), endday.Month.ToString());
            string strtem3 = string.Format("{0}年\r\n{1}-{2}月", beginday.Year.ToString(), beginday.Month.ToString(), endday.Month.ToString());
            string strtem4 = string.Format("{0}年\r\n{1}-{2}月", beginday.AddYears(-1).Year.ToString(), beginday.Month.ToString(), endday.Month.ToString());
            //导出数据到Excel
            exl.GetFirst().SetReplace("[DATE1]", strtem1 + "和" + strtem2);
            exl.GetCell(2, 2).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 3).SetReplace("[DATE3]", strtem3);
            exl.GetCell(2, 5).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 6).SetReplace("[DATE3]", strtem3);
            exl.GetCell(2, 8).SetReplace("[DATE2]", strtem4);
            exl.GetCell(2, 9).SetReplace("[DATE3]", strtem3);
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.GetRange(beginrow, 1, endrow, column).StyleLine();
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //打开
            exl.OpenExcel(temp_file_path);
        }

        /// <summary>
        /// 导出模板：每月3门急诊数据同期比表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：开始日期
        /// 3：结束日期
        /// 4：模版文件夹名称
        /// 5：模版文件名称
        /// 6：保存文件夹名称
        /// 7：保存文件名称
        /// </summary>
        public void MonthReport3()
        {
            //数据
            DataTable data = (DataTable)Params[0];
            //写入数据开始行
            int beginrow = int.Parse(Params[1].ToString());
            //开始时间
            DateTime beginday = (DateTime)Params[2];
            DateTime endday = (DateTime)Params[3];
            //模板文件夹
            string template_dir = Params[4].ToString();
            //模板文件
            string template_file = Params[5].ToString();
            //保存文件夹
            string temp_dir = Params[6].ToString();
            //保存文件
            string temp_file = Params[7].ToString();
            //模版路径
            string template_dir_path = Helper.DirPath(template_dir);
            string template_file_path = template_dir_path + template_file;
            //保存路径
            string temp_dir_path = Helper.DirPath(temp_dir);
            string temp_file_path = template_dir_path + temp_file;
            FileHelper.DirCreate(template_dir_path, false);
            FileHelper.DirCreate(temp_dir_path, false);
            //写入数据结束行
            int endrow = beginrow + data.Rows.Count - 1;
            //写入数据结束行
            int column = data.Columns.Count;
            //临时数据
            string _temp = string.Empty;
            //导入模版
            ExcelHelper exl = new ExcelHelper(template_file_path);
            //设置单元格数据
            exl.GetCell(beginrow, 1).SetCell(endday.AddYears(-1).ToString("yyyy年MM月"));
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][1].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][4].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][7].ToString());
            beginrow++;
            exl.GetCell(beginrow, 1).SetCell(endday.ToString("yyyy年MM月"));
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][2].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][5].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][8].ToString());
            beginrow++;
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][3].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][6].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][9].ToString());
            beginrow += 2;
            int iii = 9;
            exl.GetCell(beginrow, 1).SetCell(string.Format("{0}-{1}月", beginday.AddYears(-1).ToString("yyyy年MM"), endday.Month.ToString()));
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][1 + iii].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][4 + iii].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][7 + iii].ToString());
            beginrow++;
            exl.GetCell(beginrow, 1).SetCell(string.Format("{0}-{1}月", beginday.ToString("yyyy年MM"), endday.Month.ToString()));
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][2 + iii].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][5 + iii].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][8 + iii].ToString());
            beginrow++;
            exl.GetCell(beginrow, 2).SetCell(data.Rows[0][3 + iii].ToString());
            exl.GetCell(beginrow, 3).SetCell(data.Rows[0][6 + iii].ToString());
            exl.GetCell(beginrow, 4).SetCell(data.Rows[0][9 + iii].ToString());
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //打开
            exl.OpenExcel(temp_file_path);
        }


        /// <summary>
        /// 导出模板：每月4每月手术人数表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：开始日期
        /// 3：结束日期
        /// 4：模版文件夹名称
        /// 5：模版文件名称
        /// 6：保存文件夹名称
        /// 7：保存文件名称
        /// </summary>
        public void MonthReport4()
        {
            //数据
            DataTable data = (DataTable)Params[0];
            //写入数据开始行
            int beginrow = int.Parse(Params[1].ToString());
            //开始时间
            DateTime beginday = (DateTime)Params[2];
            DateTime endday = (DateTime)Params[3];
            //模板文件夹
            string template_dir = Params[4].ToString();
            //模板文件
            string template_file = Params[5].ToString();
            //保存文件夹
            string temp_dir = Params[6].ToString();
            //保存文件
            string temp_file = Params[7].ToString();
            //模版路径
            string template_dir_path = Helper.DirPath(template_dir);
            string template_file_path = template_dir_path + template_file;
            //保存路径
            string temp_dir_path = Helper.DirPath(temp_dir);
            string temp_file_path = template_dir_path + temp_file;
            FileHelper.DirCreate(template_dir_path, false);
            FileHelper.DirCreate(temp_dir_path, false);
            //写入数据结束行
            int endrow = beginrow + data.Rows.Count - 1;
            //写入数据结束行
            int column = data.Columns.Count;
            //临时数据
            string _temp = string.Empty;
            //导入模版
            ExcelHelper exl = new ExcelHelper(template_file_path);
            //设置单元格日期
            exl.GetFirst().SetReplace("[DATE1]", string.Format("{0}-{1}月", beginday.ToString("yyyy年M"), endday.Month.ToString()));
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //设置单元格计数
            exl.GetFirstRow(2).SetReplace("[NUM]", exl.GetCellToText(endrow, 6));
            //添加边框
            exl.GetRange(beginrow, 1, endrow, column).StyleLine();
            //字体红色加粗
            exl.GetRow(endrow).StyleFontColorRed().StyleFontBold();
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //打开
            exl.OpenExcel(temp_file_path);
        }
    }
}