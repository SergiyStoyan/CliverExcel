//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        //static public ICell GetCell(this IRow r, string header, bool create)
        //{
        //    ICell c = r.GetCell(x - 1);
        //    if (c == null && create)
        //        return r.CreateCell(x - 1);
        //    return c;
        //}

        static public ICell GetCell(this IRow r, int x, bool create)
        {
            ICell c = r.GetCell(x - 1);
            if (c == null && create)
                return r.CreateCell(x - 1);
            return c;
        }

        static public void Highlight(this IRow row, ICellStyle style, Excel.Color color)
        {
            row.RowStyle = Excel.highlight(row.Sheet.Workbook, style, color);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int GetLastNotEmptyColumnInRow(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            for (int x0 = row.Cells.Count - 1; x0 >= 0; x0--)
            {
                var c = row.Cells[x0];
                if (!string.IsNullOrWhiteSpace(c.GetValueAsString()))
                {
                    if (includeMerged)
                    {
                        var r = c.GetMergedRange();
                        if (r != null)
                            return r.X2;
                    }
                    return c.ColumnIndex + 1;
                }
            }
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int GetLastColumnInRow(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            if (includeMerged)
            {
                var c = row.Cells[row.Cells.Count - 1];
                var r = c.GetMergedRange();
                if (r != null)
                    return r.X2;
                return c.ColumnIndex + 1;
            }
            return row.LastCellNum;
        }

        static public IEnumerable<ICell> GetCells(this IRow row, bool createMissingInnnerCells)
        {
            if (row == null || row.Cells.Count < 1)
                yield break;
            for (int x0 = 0; x0 < row.Cells.Count; x0++)
            {
                var c = row.Cells[x0];
                if (c == null)
                {
                    if (!createMissingInnnerCells)
                        continue;
                    c = row.CreateCell(x0);
                }
                yield return c;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns>1-based</returns>
        static public int Y(this IRow row)
        {
            return row.RowNum + 1;
        }

        static public void WriteRow(this IRow row, IEnumerable<object> values)
        {
            int x = 1;
            foreach (object v in values)
                row.GetCell(x++, true).SetValue(v);
        }

        static public void WriteRow(this IRow row, params object[] values)
        {
            WriteRow(row, (IEnumerable<object>)values);
        }
    }
}