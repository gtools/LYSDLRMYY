using System;
using System.Data;
using GTSharp.Core;

namespace LYSDLRMYY
{
    /// <summary>
    /// 公共帮助类
    /// </summary>
    public class Helper
    {
        /// <summary>
        /// 获取当前工作目录的完全限定路径
        /// </summary>
        public static string CurrentDirectory { get { return Environment.CurrentDirectory; } }

        /// <summary>
        /// 拼接路径最后带(\)
        /// </summary>
        /// <param name="Path">路径</param>
        /// <param name="Dir">文件夹</param>
        /// <returns></returns>
        public static string DirPath(string Path, string Dir)
        {
            return string.Format("{0}\\{1}\\", Path.Substring(Path.Length - 1) == "\\" ? Path.Remove(Path.Length - 1) : Path, Dir.Replace("\\", ""));
        }

        /// <summary>
        /// 拼接路径最后带(\)
        /// </summary>
        /// <param name="Dir">文件夹</param>
        /// <returns></returns>
        public static string DirPath(string Dir)
        {
            return string.Format("{0}\\{1}\\", CurrentDirectory, Dir.Replace("\\", ""));
        }

        /// <summary>
        /// DataTable查询
        /// </summary>
        /// <param name="data">数据</param>
        /// <param name="filterExpression">条件</param>
        /// <param name="sort">排序字段</param>
        /// <returns></returns>
        public static DataTable DataTableSelect(DataTable data, string filterExpression, string sort)
        {
            DataTable newdt = new DataTable();
            newdt = data.Clone();
            DataRow[] drs = data.Select(filterExpression, sort);
            foreach (DataRow item in drs)
            {
                newdt.Rows.Add(item.ItemArray);
            }
            return newdt;
        }






        ///// <summary>
        ///// 文件夹路径
        ///// </summary>
        ///// <param dirname=""></param>
        ///// <returns></returns>
        //public string PathDir(string dirname)
        //{
        //    return path + "\\" + dirname.Replace("\\", "") + "\\";
        //}

        ///// <summary>
        ///// 文件路径
        ///// </summary>
        ///// <param name="dirname"></param>
        ///// <param name="filename"></param>
        ///// <returns></returns>
        //public string PathFile(string dirname, string filename)
        //{
        //    return PathDir(dirname) + filename;
        //}

        /// <summary>
        /// 清空文件夹内容
        /// </summary>
        /// <param name="path">路径</param>
        public static void DirClear(string path)
        {
            FileHelper.DirCreate(path);
            FileHelper.DirClear(path);
        }

        ///// <summary>
        ///// 打开文件目录
        ///// </summary>
        ///// <param name="dirname">路径</param>
        //public void OpenDirPath(string dirname)
        //{
        //    FileHelper.OpenDir(PathDir(dirname));
        //}



        ///// <summary>
        ///// 替换单元格数据:[DATE]替换当前时间
        ///// </summary>
        ///// <param name="addday">添加时间</param>
        ///// <param name="excel">ExcelHelper</param>
        //public void SetReplace_DATE(int addday, ExcelHelper excel)
        //{

        //    excel.SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"), excel.GetCell(1, 1));
        //}

        ///// <summary>
        ///// 替换单元格数据:[NUM]计数
        ///// </summary>
        ///// <param name="num">计数</param>
        ///// <param name="range">区域</param>
        ///// <param name="excel">ExcelHelper</param>
        //public void SetReplace_NUM(int num, Microsoft.Office.Interop.Excel.Range range, ExcelHelper excel)
        //{
        //    excel.SetReplace("[NUM]", num.ToString(), range);
        //}

        ///// <summary>
        ///// 设置字体整行红色
        ///// </summary>
        ///// <param name="rowIndex">行号</param>
        ///// <param name="excel">ExcelHelper</param>
        //public void StyleFontColorRedRow(int rowIndex, ExcelHelper excel)
        //{
        //    excel.StyleFontColor(GTSharp.ExcelColor.红色, excel.GetRow(rowIndex));
        //}

        ///// <summary>
        ///// 设置字体红色
        ///// </summary>
        ///// <param name="rowIndex">行号</param>
        ///// <param name="columnIndex">列号</param>
        ///// <param name="excel">ExcelHelper</param>
        //public void StyleFontColorRed(int rowIndex, int columnIndex, ExcelHelper excel)
        //{
        //    excel.StyleFontColor(GTSharp.ExcelColor.红色, excel.GetCell(rowIndex, columnIndex));
        //}
    }
}