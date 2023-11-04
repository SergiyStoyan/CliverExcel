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
        static public void _SetStyle(this IRow row, ICellStyle style, Excel.RowStyleMode rowStyleMode)
        {
            switch (rowStyleMode)
            {
                case Excel.RowStyleMode.RowOnly:
                    row.RowStyle = style;
                    break;
                case Excel.RowStyleMode.RowAndCells:
                    row.RowStyle = style;
                    foreach (ICell c in row.Cells)
                        c.CellStyle = style;
                    break;
                case Excel.RowStyleMode.CellsOnly:
                    foreach (ICell c in row.Cells)
                        c.CellStyle = style;
                    break;
                default:
                    throw new Exception("Unknown option: " + rowStyleMode);
            }
        }

        static public void _ShiftCellsRight(this IRow row, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            if (shift < 0)
                throw new Exception("Shift cannot be < 0: " + shift);
            for (int x = row._GetLastColumn(true); x >= x1; x--)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x + shift, onFormulaCellMoved);
        }

        static public void _ShiftCellsLeft(this IRow row, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            if (shift < 0)
                throw new Exception("Shift cannot be < 0: " + shift);
            if (shift >= x1)
                throw new Exception("Shifting left before the first column: shift=" + shift + ", x1=" + x1);
            int x2 = row._GetLastColumn(true);
            for (int x = x1; x <= x2; x++)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x - shift, onFormulaCellMoved);
        }

        static public void _ShiftCells(this IRow row, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            if (shift >= 0)
                _ShiftCellsRight(row, x1, shift, onFormulaCellMoved);
            else
                _ShiftCellsLeft(row, x1, -shift, onFormulaCellMoved);
        }

        static public ICell _GetCell(this IRow row, int x, bool createCell)
        {
            ICell c = row.GetCell(x - 1);
            if (c == null && createCell)
                return row.CreateCell(x - 1);
            return c;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyColumn(this IRow row, bool includeMerged)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            for (int i = row.Cells.Count - 1; i >= 0; i--)
            {
                var c = row.Cells[i];
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
        static public int _GetLastColumn(this IRow row, bool includeMerged)
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

        /// <summary>
        /// Get all the cells up to the last one.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="createCells"></param>
        /// <returns></returns>
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


        static public void _Clear(this IRow row, bool clearMerging)
        {
            if (clearMerging)
                row._ClearMerging();
            row.Sheet.RemoveRow(row);
        }

        static public void _ClearMerging(this IRow row)
        {
            new Excel.Range(row.Sheet, row._Y(), 1, row._Y(), null).ClearMerging();
        }

        /// <summary> 
        /// Value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        static public object _GetValue(this IRow row, int x)
        {
            return row._GetCell(x, false)?._GetValue();
        }

        /// <summary> 
        /// Set value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        static public void _SetValue(this IRow row, int x, object value)
        {
            row._GetCell(x, false)?._SetValue(value);
        }

        /// <summary>
        /// Value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="x"></param>
        /// <param name="allowNull"></param>
        /// <returns></returns>
        static public string _GetValueAsString(this IRow row, int x, bool allowNull = false)
        {
            ICell c = row._GetCell(x, false);
            if (c == null)
                return allowNull ? null : string.Empty;
            return c._GetValueAsString(allowNull);
        }

        /// <summary>
        /// Images anchored in the specified cell coordinates. The cell may not exist.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        static public IEnumerable<Excel.Image> _GetImages(this IRow row, int x)
        {
            return row.Sheet._GetImages(row._Y(), x);
        }

        static public Uri _GetLink(this IRow row, int x)
        {
            return row._GetCell(x, false)?._GetLink();
        }

        static public void _SetLink(this IRow row, int x, Uri uri)
        {
            row._GetCell(x, true)._SetLink(uri);
        }
    }
}