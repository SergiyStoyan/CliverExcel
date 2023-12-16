//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using static Cliver.Excel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public void _SetAlteredStyles<T>(this IRow row, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
        {
            var styleCache = row.Sheet.Workbook._Excel().OneWorkbookStyleCache;
            foreach (ICell cell in row.Cells)
                cell.CellStyle = styleCache.GetAlteredStyle(cell.CellStyle, alterationKey, alterStyle, reuseUnusedStyle);
        }

        static public void _Copy(this IRow row, int y2, CopyCellMode copyCellMode = null)
        {
            //if (row == null)
            //    return;            
            if (row._Y() == y2)
                return;
            IRow r2 = row.Sheet._GetRow(y2 - 1, true);
            r2._Clear(false);
            r2.Height = row.Height;
            foreach (ICell c1 in row.Cells)
                c1._Copy(y2, c1._X(), copyCellMode);
        }

        /// <summary>
        /// Insert a copy and remove the source.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="y2"></param>
        /// <param name="moveRegionMode"></param>
        static public void _Move(this IRow row, int y2, MoveRegionMode moveRegionMode = null)
        {
            row.Sheet._ShiftRowsDown(y2, 1, moveRegionMode);

            if (moveRegionMode?.UpdateMergedRegions == true)
            {
                row.Sheet.MergedRegions.ForEach(a =>
                {
                    if (a.FirstRow < y2 - 1)
                    {
                        if (a.LastRow >= y2 - 1)
                            a.LastRow += 1;
                    }
                    else
                    {
                        a.FirstRow += 1;
                        a.LastRow += 1;
                    }
                });
            }

            row._Copy(y2, moveRegionMode);
            row._Remove(moveRegionMode);
        }

        /// <summary>
        /// Remove the row from its sheet and (!)shift rows below which can be slow. Not to shift, use Clear()
        /// </summary>
        /// <param name="row"></param>
        /// <param name="moveRegionMode"></param>
        /// <param name="preserveCells">
        /// (!)Done in a hacky way through Reflection so might change with POI update.
        /// (!)GetCell() might work incorrectly on such rows.
        /// </param>
        static public void _Remove(this IRow row, MoveRegionMode moveRegionMode = null, bool preserveCells = false)
        {
            if (moveRegionMode?.UpdateMergedRegions == true)
            {
                for (int i = row.Sheet.MergedRegions.Count - 1; i >= 0; i--)
                {
                    NPOI.SS.Util.CellRangeAddress a = row.Sheet.GetMergedRegion(i);
                    if (a.FirstRow < row.RowNum)
                    {
                        if (a.LastRow >= row.RowNum)
                            a.LastRow -= 1;
                    }
                    else if (a.FirstRow == row.RowNum && a.LastRow == row.RowNum)
                        row.Sheet.RemoveMergedRegion(i);
                    else
                    {
                        a.FirstRow -= 1;
                        a.LastRow -= 1;
                    }
                }
            }

            if (preserveCells
                && row is XSSFRow xSSFRow//in HSSFRow Cells remain after removing the row
                )
            {
                SortedDictionary<int, ICell> cells = new SortedDictionary<int, ICell>();
                row.Cells.Select((a, i) => (a, i)).ForEach(a => cells.Add(a.i, a.a));
                if (XSSFRow_cells_FI == null)
                    XSSFRow_cells_FI = xSSFRow.GetType().GetField("_cells", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                //cells = (SortedDictionary<int, ICell>)XSSFRow_cells_FI.GetValue(xSSFRow);!!!will be NULLed when removing

                row.Sheet.RemoveRow(row);

                XSSFRow_cells_FI.SetValue(row, cells);
            }
            else
                row.Sheet.RemoveRow(row);

            row.Sheet._ShiftRowsUp(row.RowNum + 2, 1, moveRegionMode);
        }
        static System.Reflection.FieldInfo XSSFRow_cells_FI = null;

        public static int _LastCellX(this IRow row)
        {
            return row.LastCellNum + 1;
        }

        static public void _RemoveCell(this IRow row, int x)
        {
            var c = row.GetCell(x - 1);
            if (c != null)
                row.RemoveCell(c);
        }

        static public void _MoveCell(this IRow row, int x1, int x2)
        {
            var c = row.GetCell(x1 - 1);
            if (c != null)
                row.MoveCell(c, x2 - 1);
        }

        //static public void _SetStyle(this IRow row, ICellStyle style, RowStyleMode rowStyleMode)
        //{
        //    if (rowStyleMode.HasFlag(RowStyleMode.Row))
        //        row.RowStyle = style;
        //    if (rowStyleMode.HasFlag(RowStyleMode.ExistingCells))
        //        foreach (ICell c in row.Cells)
        //            c.CellStyle = style;
        //    else if (rowStyleMode.HasFlag(RowStyleMode.NoGapCells))
        //        for (int x = row.LastCellNum; x > 0; x--)
        //            row._GetCell(x, true).CellStyle = style;
        //}

        static public void _ShiftCellsRight(this IRow row, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            if (shift < 0)
                throw new Exception("Shift cannot be < 0: " + shift);
            for (int x = row._GetLastColumn(true); x >= x1; x--)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x + shift, copyCellMode);
        }

        static public void _ShiftCellsLeft(this IRow row, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            if (shift < 0)
                throw new Exception("Shift cannot be < 0: " + shift);
            if (shift >= x1)
                throw new Exception("Shifting left before the first column: shift=" + shift + ", x1=" + x1);
            int x2 = row._GetLastColumn(true) + 1;
            for (int x = x1; x <= x2; x++)
                row.Sheet._MoveCell(row._Y(), x, row._Y(), x - shift, copyCellMode);
        }

        static public void _ShiftCells(this IRow row, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            if (shift >= 0)
                _ShiftCellsRight(row, x1, shift, copyCellMode);
            else
                _ShiftCellsLeft(row, x1, -shift, copyCellMode);
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

        /// <summary>
        /// Delete the row as an object but not shift rows below.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="clearMerging"></param>
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

        static public string _GetLink(this IRow row, int x)
        {
            return row?._GetCell(x, false)?._GetLink();
        }

        static public void _SetLink(this IRow row, int x, string link)
        {
            row._GetCell(x, true)._SetLink(link);
        }
    }
}