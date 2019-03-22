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
        /// 路径
        /// </summary>
        string path { get; set; }
        /// <summary>
        /// 构造函数
        /// </summary>
        public Helper()
        {
            path = Environment.CurrentDirectory;
        }

        /// <summary>
        /// 文件夹路径
        /// </summary>
        /// <param dirname=""></param>
        /// <returns></returns>
        public string PathDir(string dirname)
        {
            return path + "\\" + dirname.Replace("\\", "") + "\\";
        }

        /// <summary>
        /// 文件路径
        /// </summary>
        /// <param name="dirname"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string PathFile(string dirname, string filename)
        {
            return PathDir(dirname) + filename;
        }

        /// <summary>
        /// 清空文件夹内容
        /// </summary>
        /// <param name="dirname">路径</param>
        public void DirClear(string dirname)
        {
            FileHelper.DirCreate(PathDir(dirname));
            FileHelper.DirClear(PathDir(dirname));
        }

        /// <summary>
        /// 打开文件目录
        /// </summary>
        /// <param name="dirname">路径</param>
        public void OpenDirPath(string dirname)
        {
            FileHelper.OpenDir(PathDir(dirname));
        }

        /// <summary>
        /// DataTable查询
        /// </summary>
        /// <param name="data">数据</param>
        /// <param name="filterExpression">条件</param>
        /// <param name="sort">排序字段</param>
        /// <returns></returns>
        public DataTable DataTableSelect(DataTable data, string filterExpression, string sort)
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

        /// <summary>
        /// 替换单元格数据:[DATE]替换当前时间
        /// </summary>
        /// <param name="addday">添加时间</param>
        /// <param name="excel">ExcelHelper</param>
        public void SetReplace_DATE(int addday, ExcelHelper excel)
        {
            excel.SetReplace("[DATE]", DateTime.Now.AddDays(addday).ToString("yyyy年MM月dd日"), excel.GetCell(1, 1));
        }

        /// <summary>
        /// 替换单元格数据:[NUM]计数
        /// </summary>
        /// <param name="num">计数</param>
        /// <param name="range">区域</param>
        /// <param name="excel">ExcelHelper</param>
        public void SetReplace_NUM(int num, Microsoft.Office.Interop.Excel.Range range, ExcelHelper excel)
        {
            excel.SetReplace("[NUM]", num.ToString(), range);
        }

        /// <summary>
        /// 设置字体整行红色
        /// </summary>
        /// <param name="rowIndex">行号</param>
        /// <param name="excel">ExcelHelper</param>
        public void StyleFontColorRedRow(int rowIndex, ExcelHelper excel)
        {
            excel.StyleFontColor(GTSharp.ExcelColor.红色, excel.GetRow(rowIndex));
        }

        /// <summary>
        /// 设置字体红色
        /// </summary>
        /// <param name="rowIndex">行号</param>
        /// <param name="columnIndex">列号</param>
        /// <param name="excel">ExcelHelper</param>
        public void StyleFontColorRed(int rowIndex, int columnIndex, ExcelHelper excel)
        {
            excel.StyleFontColor(GTSharp.ExcelColor.红色, excel.GetCell(rowIndex, columnIndex));
        }
    }
}