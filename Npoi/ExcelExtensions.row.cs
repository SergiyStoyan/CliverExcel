//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public void _ShiftCellsRight(this IRow row, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int x = row._GetLastColumn(true); x >= x1; x--)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x + shift, onFormulaCellMoved);
        }

        static public void _ShiftCellsLeft(this IRow row, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            int x2 = row._GetLastColumn(true);
            for (int x = x1; x <= x2; x++)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x - shift, onFormulaCellMoved);
        }

        //static public ICell GetCell(this IRow r, string header, bool create)
        //{
        //    ICell c = r.GetCell(x - 1);
        //    if (c == null && create)
        //        return r.CreateCell(x - 1);
        //    return c;
        //}

        static public ICell _GetCell(this IRow r, int x, bool createCell)
        {
            ICell c = r.GetCell(x - 1);
            if (c == null && createCell)
                return r.CreateCell(x - 1);
            return c;
        }

        //static public void Highlight(this IRow row, ICellStyle style, Excel.Color color)
        //{
        //    row.RowStyle = Excel.highlight(row.Sheet.Workbook, style, color);
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyColumn(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            for (int x0 = row.Cells.Count - 1; x0 >= 0; x0--)
            {
                var c = row.GetCell(x0);
                if (!string.IsNullOrWhiteSpace(c?._GetValueAsString()))
                {
                    if (includeMerged)
                    {
                        var r = c._GetMergedRange();
                        if (r != null)
                            return r.X2.Value;
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
        static public int _GetLastColumn(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            if (includeMerged)
            {
                var c = row.Cells[row.Cells.Count - 1];
                var r = c._GetMergedRange();
                if (r != null)
                    return r.X2.Value;
                return c.ColumnIndex + 1;
            }
            return row.LastCellNum;
        }

        static public IEnumerable<ICell> _GetCells(this IRow row, bool createCells)
        {
            return _GetCellsInRange(row, createCells);
        }

        static public IEnumerable<ICell> _GetCellsInRange(this IRow row, bool createCells, int x1 = 1, int? x2 = null)
        {
            if (row == null)
                yield break;
            if (x2 == null)
                x2 = row.LastCellNum;
            for (int x = x1; x <= x2; x++)
                yield return row._GetCell(x, createCells);
        }

        /// <summary>
        /// 1-based row index on the sheet.
        /// </summary>
        /// <param name="row"></param>
        /// <returns>1-based</returns>
        static public int _Y(this IRow row)
        {
            return row.RowNum + 1;
        }

        static public void _Write<T>(this IRow row, IEnumerable<T> values)
        {
            int x = 1;
            foreach (T v in values)
                row._GetCell(x++, true)._SetValue(v);
        }

        static public void _Write(this IRow row, params string[] values)
        {
            _Write(row, (IEnumerable<string>)values);
        }

        static public void _SetStyles(this IRow row, int x1, IEnumerable<ICellStyle> styles)
        {
            _SetStyles(row, x1, styles.ToArray());
        }

        static public void _SetStyles(this IRow row, int x1, params ICellStyle[] styles)
        {
            var cs = row._GetCellsInRange(true, x1, styles.Length).ToList();
            for (int i = x1 - 1; i < styles.Length; i++)
                cs[i].CellStyle = styles[i];
        }


        static public void _Clear(this IRow row, int y, bool clearMerging)
        {
            if (clearMerging)
                row._ClearMerging();
            row.Sheet.RemoveRow(row);
        }

        static public void _ClearMerging(this IRow row)
        {
            new Excel.Range(row.Sheet, row._Y(), 1, row._Y(), null).ClearMerging();
        }
    }
}