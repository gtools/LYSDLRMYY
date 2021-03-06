﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using GTSharp;
using GTSharp.Core;

namespace LYSDLRMYY
{
    /// <summary>
    /// 每日院长查询报表
    /// </summary>
    public class MRYZCXBB
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
        public MRYZCXBB()
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
        /// 导出模板：每日1科室在院人数一览表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void DailyReport1()
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
            exl.GetFirst().SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"));
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.GetRange(beginrow, 1, endrow, column).StyleLine();
            //添加字体红色
            exl.GetRow(endrow).StyleFontColorRed();
            //合计为0删除列
            for (int i = data.Columns.Count; i >= 1; i--)
            {
                //获取单元格数据
                _temp = exl.GetCellToText(endrow, i);
                //删除列
                if (_temp == "0" || _temp.IsNullOrWhiteSpace())
                    exl.DelColumn(i);
            }
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //保存并打开
            exl.OpenExcel(temp_file_path);
        }

        /// <summary>
        /// 导出模板：每日2按手术时间统计手术人数表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void DailyReport2()
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
            exl.GetFirst().SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"));
            //设置单元格计数
            exl.GetFirstRow(2).SetReplace("[NUM]", data.Rows.Count.ToString());
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.GetRange(beginrow, 1, endrow, column).StyleLine();
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //保存并打开
            exl.OpenExcel(temp_file_path);
        }

        /// <summary>
        /// 导出模板：每日3在院危重病人患者明细表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void DailyReport3()
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
            GTSharp.Core.ExcelHelper exl = new GTSharp.Core.ExcelHelper(template_file_path);
            //设置单元格日期
            exl.GetFirst().SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"));
            //设置单元格计数
            exl.GetFirstRow(2).SetReplace("[NUM]", data.Rows.Count.ToString());
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.StyleLine(exl.GetRange(beginrow, 1, endrow, column));
            //判断日期字体变红
            for (int i = 0; i < data.Rows.Count; i++)
            {
                _temp = exl.GetCellToText(i + beginrow, column);
                if (_temp == DateTime.Now.ToString("yyyy-MM-dd") || _temp == DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"))
                    exl.GetRow(i + beginrow).StyleFontColorRed();
            }
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //保存并打开
            exl.OpenExcel(temp_file_path);
        }
        
        /// <summary>
        /// 导出模板：每日4在院I级护理患者明细表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void DailyReport4()
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
            GTSharp.Core.ExcelHelper exl = new GTSharp.Core.ExcelHelper(template_file_path);
            //设置单元格日期
            exl.GetFirst().SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"));
            //设置单元格计数
            exl.GetFirstRow(2).SetReplace("[NUM]", data.Rows.Count.ToString());
            //导出数据到Excel
            exl.DataTableToExcel(data, beginrow);
            //添加边框
            exl.StyleLine(exl.GetRange(beginrow, 1, endrow, column));
            //判断日期字体变红
            for (int i = 0; i < data.Rows.Count; i++)
            {
                _temp = exl.GetCellToText(i + beginrow, column);
                if (_temp == DateTime.Now.ToString("yyyy-MM-dd") || _temp == DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"))
                    exl.GetRow(i + beginrow).StyleFontColorRed();
            }
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //保存并打开
            exl.OpenExcel(temp_file_path);
        }
        
        /// <summary>
        /// 导出模板：每日5主要业务数据表.xls
        /// 参数
        /// 0：DataTable数据
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模版文件夹名称
        /// 4：模版文件名称
        /// 5：保存文件夹名称
        /// 6：保存文件名称
        /// </summary>
        public void DailyReport5()
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
            GTSharp.Core.ExcelHelper exl = new GTSharp.Core.ExcelHelper(template_file_path);
            //设置单元格日期
            exl.GetFirst().SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"));
            //设置单元格数据
            //全院收入
            exl.SetCell(5, 6, data.Rows[0][0]);
            //住院收入
            exl.SetCell(4, 6, data.Rows[0][1]);
            //门诊收入
            exl.SetCell(3, 6, data.Rows[0][2]);
            //全院药品收入
            exl.SetCell(5, 2, data.Rows[0][3]);
            //住院药品收入
            exl.SetCell(4, 2, data.Rows[0][4]);
            //门诊药品收入
            exl.SetCell(3, 2, data.Rows[0][5]);
            //全院药占比
            exl.SetCell(5, 3, data.Rows[0][6]);
            //住院药占比
            exl.SetCell(4, 3, data.Rows[0][7]);
            //门诊药占比
            exl.SetCell(3, 3, data.Rows[0][8]);
            //全院人次
            //exl.SetCell(_dtstartheight, 2, data.Rows[0][9]);
            //住院人次
            exl.SetCell(4, 4, data.Rows[0][10]);
            //门诊人次
            exl.SetCell(3, 4, data.Rows[0][11]);
            //全院平均
            //exl.SetCell(_dtstartheight, 2, _dt.Rows[0][12]);
            //住院平均
            exl.SetCell(4, 5, data.Rows[0][13]);
            //门诊平均
            exl.SetCell(3, 5, data.Rows[0][14]);
            //保存
            exl.SaveAsFile(temp_file_path);
            Thread.Sleep(50);
            //保存并打开
            exl.OpenExcel(temp_file_path);
        }
    }
}