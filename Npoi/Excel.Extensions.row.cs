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
        static public int GetLastNotEmptyColumn(this IRow row, bool includeMerged = true)
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
        static public int GetLastColumn(this IRow row, bool includeMerged = true)
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

        static public IEnumerable<ICell> GetCells(this IRow row, bool createCells)
        {
            return GetCellsInRange(row, createCells);
        }

        static public IEnumerable<ICell> GetCellsInRange(this IRow row, bool createCells, int y1 = 1, int? y2 = null)
        {
            if (row == null)
                yield break;
            if (y2 == null)
                y2 = row.LastCellNum;
            for (int x = y1; x <= y2; x++)
                yield return row.GetCell(x, createCells);
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

        static public void Write(this IRow row, IEnumerable<object> values)
        {
            int x = 1;
            foreach (object v in values)
                row.GetCell(x++, true).SetValue(v);
        }

        static public void Write(this IRow row, params object[] values)
        {
            Write(row, (IEnumerable<object>)values);
        }

        static public void SetStyles(this IRow row, int y1, IEnumerable<ICellStyle> styles)
        {
            SetStyles(row, y1, styles.ToArray());
        }

        static public void SetStyles(this IRow row, int y1, params ICellStyle[] styles)
        {
            var cs = row.GetCellsInRange(true, y1, styles.Length).ToList();
            for (int i = y1 - 1; i < styles.Length; i++)
                cs[i].CellStyle = styles[i];
        }
    }
}