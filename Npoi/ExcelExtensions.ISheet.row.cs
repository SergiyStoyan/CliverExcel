//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************

using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using static Cliver.Excel;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        /// <summary>
        /// Last row existing as an object. It still can be empty as the previous ones.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static int _LastRowY(this ISheet sheet)
        {
            return sheet.LastRowNum + 1;
        }

        /// <summary>
        /// Removes empty rows.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="includeEmptyCellRows">might considerably slow down if TRUE</param>
        /// <param name="shiftRowsBelow">the remaining rows are shifted if TRUE</param>
        public static void _RemoveEmptyRows(this ISheet sheet, bool includeEmptyCellRows, bool shiftRowsBelow)
        {
            int removedRowsCount = 0;
            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                var r = sheet.GetRow(i);
                if (r != null)
                {
                    if (r.LastCellNum >= 0
                        && (!includeEmptyCellRows || r._GetLastNotEmptyColumn(false) > 0)
                        )
                    {
                        if (removedRowsCount > 0)
                        {
                            if (r.RowNum + 1 + removedRowsCount <= sheet.LastRowNum)
                                sheet.ShiftRows(r.RowNum + 1 + removedRowsCount, sheet.LastRowNum, -removedRowsCount);
                            removedRowsCount = 0;
                        }
                        continue;
                    }
                    sheet.RemoveRow(r);
                }
                if (shiftRowsBelow)
                    removedRowsCount++;
            }
            if (removedRowsCount > 0
                && 1 + removedRowsCount <= sheet.LastRowNum
                )
                sheet.ShiftRows(1 + removedRowsCount, sheet.LastRowNum, -removedRowsCount);
        }

        static public IEnumerable<IRow> _GetRows(this ISheet sheet, RowScope rowScope)
        {
            return sheet._GetRowsInRange(rowScope);
        }

        static public int _GetLastRow(this ISheet sheet, LastRowCondition lastRowCondition, bool includeMerged)
        {
            IRow row = null;
            switch (lastRowCondition)
            {
                case LastRowCondition.NotEmpty:
                    return sheet._GetLastNotEmptyRow(includeMerged);
                case LastRowCondition.HasCells:
                    for (int i = sheet.LastRowNum; i >= 0; i--)
                    {
                        row = sheet.GetRow(i);
                        if (row == null)
                            continue;
                        if (row.LastCellNum >= 0)
                            break;
                    }
                    break;
                case LastRowCondition.NotNull:
                    row = sheet.GetRow(sheet.LastRowNum);
                    break;
                default:
                    throw new Exception("Unknown option: " + lastRowCondition.ToString());
            }
            if (row == null)
                return 0;
            if (!includeMerged)
                return row._Y();
            int maxY = 0;
            foreach (var mr in sheet.MergedRegions)
                foreach (var c in row.Cells)
                    if (mr.IsInRange(c.RowIndex, c.ColumnIndex))
                        if (maxY < mr.LastRow)
                            maxY = mr.LastRow;
            return maxY;
        }

        static public IRow _GetRow(this ISheet sheet, int y, bool createRow)
        {
            IRow r = sheet.GetRow(y - 1);
            if (r == null && createRow)
                r = sheet.CreateRow(y - 1);
            return r;
        }

        /// <summary>
        /// (!)May return a huge pile of null and empty rows after the last actual row. 
        /// Use RowScope.ExistingOnly to avoid garbage or call _RemoveEmptyRows() before.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowScope"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        /// <returns></returns>
        static public IEnumerable<IRow> _GetRowsInRange(this ISheet sheet, RowScope rowScope, int y1 = 1, int? y2 = null)
        {
            if (y2 == null)
                y2 = sheet.LastRowNum + 1;
            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
            switch (rowScope)
            {
                case RowScope.CreateIfNull:
                    for (int i = y1 - 1; i < y2; i++)
                    {
                        var r = sheet.GetRow(i);
                        if (r == null)
                            r = sheet.CreateRow(i);
                        yield return r;
                    }
                    break;
                case RowScope.IncludeNull:
                    for (int i = y1 - 1; i < y2; i++)
                    {
                        var r = sheet.GetRow(i);
                        yield return r;
                    }
                    break;
                case RowScope.NotNull:
                    for (int i = y1 - 1; i < y2; i++)
                    {
                        var r = sheet.GetRow(i);
                        if (r != null)
                            yield return r;
                    }
                    break;
                case RowScope.WithCells:
                    for (int i = y1 - 1; i < y2; i++)
                    {
                        var r = sheet.GetRow(i);
                        if (r != null && r.LastCellNum >= 0)
                            yield return r;
                    }
                    break;
                case RowScope.NotEmpty:
                    for (int i = y1 - 1; i < y2; i++)
                    {
                        var r = sheet.GetRow(i);
                        if (r != null && r._GetLastNotEmptyColumn(false) > 0)
                            yield return r;
                    }
                    break;
                default:
                    throw new Exception("Unknown option: " + rowScope.ToString());
            }
        }

        /// <summary>
        /// (!)It does not care about formulas and links. Shift*() does.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        static public IRow _AppendRow<T>(this ISheet sheet, IEnumerable<T> values)
        {
            int lastRowY = sheet._GetLastRow(LastRowCondition.HasCells, false);
            return sheet._WriteRow(lastRowY + 1, values);
        }

        /// <summary>
        /// (!)It does not care about formulas and links. Shift*() does.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        static public IRow _AppendRow<T>(this ISheet sheet, params T[] values)
        {
            return sheet._AppendRow((IEnumerable<T>)values);
        }

        /// <summary>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="y"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        static public IRow _InsertRow<T>(this ISheet sheet, int y, IEnumerable<T> values = null, MoveRegionMode moveRegionMode = null)
        {
            int lastRowY = sheet._GetLastRow(LastRowCondition.HasCells, false);
            if (y <= lastRowY)
                sheet._ShiftRowsDown(y, 1, moveRegionMode);
            return sheet._WriteRow(y, values);
        }

        /// <summary>
        /// (!)It does not care about formulas and links. Shift*() does.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        static public IRow _InsertRow<T>(this ISheet sheet, int y, params T[] values)
        {
            return sheet._InsertRow(y, (IEnumerable<T>)values);
        }

        static public IRow _WriteRow<T>(this ISheet sheet, int? y, bool insert, IEnumerable<T> values)
        {
            return y == null ? sheet._AppendRow(values) : insert ? sheet._InsertRow(y.Value, values) : sheet._WriteRow(y.Value, values);
        }

        static public IRow _WriteRow(this ISheet sheet, int? y, bool insert, params string[] values)
        {
            return sheet._WriteRow(y, insert, (IEnumerable<string>)values);
        }

        static public IRow _WriteRow<T>(this ISheet sheet, int y, IEnumerable<T> values)
        {
            IRow r = sheet._GetRow(y, true);
            r._Write(values);
            return r;
        }

        static public IRow _WriteRow(this ISheet sheet, int y, params string[] values)
        {
            return sheet._WriteRow(y, (IEnumerable<string>)values);
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
        static public IRow _RemoveRow(this ISheet sheet, int y, MoveRegionMode moveRegionMode = null, bool preserveCells = false)
        {
            IRow row = sheet.GetRow(y - 1);
            if (row == null)
            {
                if (moveRegionMode?.CopyComment == true)
                {//!!!TO DO remove comments                    
                    //sheet.GetCellComment()
                }
                return null;
            }

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

            return row;
        }
        static System.Reflection.FieldInfo XSSFRow_cells_FI = null;

        static public void _RemoveRowRange(this ISheet sheet, int y1, int y2, MoveRegionMode moveRegionMode = null)
        {
            for (int y = y1; y <= y2; y++)
                sheet._RemoveRow(y, moveRegionMode);
        }

        /// <summary>
        /// Delete the row as an object but not shift rows below.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y"></param>
        /// <param name="clearMerging"></param>
        static public void _ClearRow(this ISheet sheet, int y, bool clearMerging)
        {
            var r = sheet._GetRow(y, false);
            if (r != null)
                r._Clear(clearMerging);
            else
                sheet._ClearMergingInRow(y);
        }

        static public void _ClearRowRange(this ISheet sheet, int y1, int y2, bool clearMerging)
        {
            for (int y = y1; y <= y2; y++)
                sheet._ClearRow(y, clearMerging);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        //static public IRow _MoveRow(this ISheet sheet, int y1, int y2, bool insert, MoveRegionMode moveRegionMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        //{
        //    var row1 = sheet._GetRow(y1, false);
        //    if (insert)
        //    {
        //        row1.Sheet._ShiftRowsDown(y2, 1, moveRegionMode);

        //        if (moveRegionMode?.UpdateMergedRegions == true)
        //        {
        //            row1.Sheet.MergedRegions.ForEach(a =>
        //            {
        //                if (a.FirstRow < y2 - 1)
        //                {
        //                    if (a.LastRow >= y2 - 1)
        //                        a.LastRow += 1;
        //                }
        //                else
        //                {
        //                    a.FirstRow += 1;
        //                    a.LastRow += 1;
        //                }
        //            });
        //        }
        //    }
        //    else
        //        row1.Sheet._RemoveRow(y2, moveRegionMode);
        //    IRow row2 = sheet._CopyRow(row1._Y(), y2, moveRegionMode, sheet2, styleMap);
        //    row1._Remove(moveRegionMode);
        //    return row2;
        //}
        static public IRow _MoveRow(this ISheet sheet, int y1, int y2, bool insert, MoveRegionMode moveRegionMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            sheet2 = sheet2 ?? sheet;
            IRow r1 = sheet._GetRow(y1, false);
            if (sheet2 == sheet && y1 == y2)
                return r1;
            if (insert)
                sheet2._ShiftRowsDown(y2, 1, moveRegionMode);
            IRow r2 = sheet._CopyRow(r1._Y(), y2, moveRegionMode, sheet2, styleMap);
            r1._Remove(moveRegionMode);
            return r2;
        }

        static public IRow _CopyRow(this ISheet sheet, int y1, int y2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
        {
            sheet2 = sheet2 ?? sheet;
            IRow r1 = sheet._GetRow(y1, false);
            if (sheet2 == sheet && y1 == y2)
                return r1;
            sheet2._ClearRow(y2, false);
            if (r1 == null)
                return null;
            IRow r2 = sheet2._GetRow(y2, true);
            r2.Height = r1.Height;
            foreach (ICell c1 in r1.Cells)
                c1._Copy(y2, c1._X(), copyCellMode, sheet2, styleMap);
            return r2;
        }

        /// <summary>
        /// Based on ISheet.CopyRow()
        /// (!)It seems to be slower than _CopyRow()
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        static public void _CopyRow2(this ISheet sheet, int y1, int y2)
        {
            if (y1 == y2)
                return;
            var r1 = sheet._GetRow(y1, false);
            if (r1 == null)
            {
                sheet._RemoveRow(y2);
                return;
            }
            sheet.CopyRow(y1 - 1, y2 - 1);
            sheet._RemoveRow(y2);
        }

        /// <summary>
        /// Based on ISheet.CopyRow()
        /// (!)It seems to be slower than _MoveRow()
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        static public void _MoveRow2(this ISheet sheet, int y1, int y2)
        {
            sheet._CopyRow2(y1, y2);
            sheet._RemoveRow(y1);
        }

        /// <summary>
        /// Based on ISheet.ShiftRows(). 
        /// (!)On big sheets it is slower than _MoveRow()
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        static public void _MoveRow3(this ISheet sheet, int y1, int y2)
        {
            if (y1 == y2)
                return;
            if (y1 > y2)
            {
                sheet.ShiftRows(y2 - 1, sheet.LastRowNum, 1);
                sheet.ShiftRows(y1, y1, y2 - y1 - 1);
                sheet.ShiftRows(y1 + 1, sheet.LastRowNum, -1);
            }
            else
            {
                if (y2 - 1 < sheet.LastRowNum)
                    sheet.ShiftRows(y2, sheet.LastRowNum, 1);
                sheet.ShiftRows(y1 - 1, y1 - 1, y2 - y1 + 1);
                sheet.ShiftRows(y1, sheet.LastRowNum, -1);
            }
        }

        static public void _ShiftRowsDown(this ISheet sheet, int y, int shift, MoveRegionMode moveRegionMode = null)
        {
            for (int y0 = sheet.LastRowNum; y0 >= y - 1; y0--)
            {
                IRow r2 = sheet.GetRow(y0 + shift);
                IRow r1 = sheet.GetRow(y0);
                r2?._Remove();
                if (r1 == null)
                    continue;
                r2 = sheet.CreateRow(y0 + shift);
                r2.Height = r1.Height;
                foreach (ICell c1 in r1.Cells)
                    c1._Copy(c1._Y() + shift, c1._X(), moveRegionMode);
                sheet.RemoveRow(r1);
            }

            if (moveRegionMode?.UpdateMergedRegions == true)
                for (int i = sheet.MergedRegions.Count - 1; i >= 0; i--)
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(i);
                    if (a.FirstRow < y - 1)
                    {
                        if (a.LastRow < y - 1)
                        { }
                        else
                            a.LastRow += shift;
                    }
                    else
                    {
                        a.FirstColumn += shift;
                        a.LastColumn += shift;
                    }
                }
        }

        static public void _ShiftRowsUp(this ISheet sheet, int y, int shift, MoveRegionMode moveRegionMode = null)
        {
            for (int y0 = y - 1; y0 <= sheet.LastRowNum; y0++)
            {
                IRow r2 = sheet.GetRow(y0 - shift);
                IRow r1 = sheet.GetRow(y0);
                r2?._Remove();
                if (r1 == null)
                    continue;
                r2 = sheet.CreateRow(y0 - shift);
                r2.Height = r1.Height;
                foreach (ICell c1 in r1.Cells)
                    c1._Copy(c1._Y() - shift, c1._X(), moveRegionMode);
                sheet.RemoveRow(r1);
            }

            if (moveRegionMode?.UpdateMergedRegions == true)
                for (int i = sheet.MergedRegions.Count - 1; i >= 0; i--)
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(i);
                    if (a.FirstRow < y - 1)
                    {
                        if (a.LastRow < y - 1)
                        { }
                        else if (a.LastRow <= y - 1 + shift)
                            a.LastRow = y - 1;
                        else
                            a.LastRow -= shift;
                        a.LastRow -= 1;
                    }
                    else if (a.LastRow <= y - 1 + shift)
                        sheet.RemoveMergedRegion(i);
                    else
                    {
                        a.FirstRow -= shift;
                        a.LastRow -= shift;
                    }
                }
        }

        static public void _ShiftRowCellsRight(this ISheet sheet, int y, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetRow(y, false)?._ShiftCellsRight(x1, shift, copyCellMode);
        }

        static public void _ShiftRowCellsLeft(this ISheet sheet, int y, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetRow(y, false)?._ShiftCellsLeft(x1, shift, copyCellMode);
        }

        static public void _ShiftRowCells(this ISheet sheet, int y, int x1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetRow(y, false)?._ShiftCells(x1, shift, copyCellMode);
        }

        static public void _SetStyleInRow(this ISheet sheet, ICellStyle style, bool createCells, int y)
        {
            sheet._SetStyleInRowRange(style, createCells, y, y);
        }

        static public void _SetStyleInRowRange(this ISheet sheet, ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            sheet._GetRange(y1, 1, y2, null).SetStyle(style, createCells);
        }

        static public void _ReplaceStyleInRowRange(this ISheet sheet, ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            sheet._GetRange(y1, 1, y2, null).ReplaceStyle(style1, style2);
        }

        static public void _ClearStyleInRowRange(this ISheet sheet, ICellStyle style, int y1, int? y2 = null)
        {
            sheet._ReplaceStyleInRowRange(style, null, y1, y2);
        }

        static public void _AutosizeRowsInRange(this ISheet sheet, int y1 = 1, int? y2 = null)
        {
            sheet._GetRowsInRange(RowScope.NotNull, y1, y2).ForEach(a => a.Height = -1);
        }

        static public void _AutosizeRows(this ISheet sheet)
        {
            sheet._AutosizeRowsInRange();
        }

        static public void _ClearMergingInRow(this ISheet sheet, int y)
        {
            sheet._GetRange(y, 1, y, null).ClearMerging();
        }

        static public int _GetLastColumnInRowRange(this ISheet sheet, bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return sheet._GetRowsInRange(RowScope.NotNull, y1, y2).Max(a => a._GetLastColumn(includeMerged));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyColumnInRow(this ISheet sheet, bool includeMerged, int y)
        {
            IRow row = sheet._GetRow(y, false);
            if (row == null)
                return 0;
            return row._GetLastNotEmptyColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastColumnInRow(this ISheet sheet, bool includeMerged, int y)
        {
            IRow row = sheet._GetRow(y, false);
            if (row == null)
                return 0;
            return row._GetLastColumn(includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyColumnInRowRange(this ISheet sheet, bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return sheet._GetRowsInRange(RowScope.NotNull, y1, y2).Max(a => a._GetLastNotEmptyColumn(includeMerged));
        }

        /// <summary>
        /// Removes empty rows from the bottom
        /// </summary>
        /// <param name="sheet"></param>
        static public void _TrimBottomRows(this ISheet sheet)
        {
            //int lastNotEmptyY = sheet._GetLastNotEmptyRow(true);
            //for (int i = sheet.LastRowNum; i >= lastNotEmptyY; i--)
            //{
            //    var r = sheet.GetRow(i);
            //    if (r != null)
            //        sheet.RemoveRow(r);
            //}

            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                var r = sheet.GetRow(i);
                if (r == null)
                    continue;
                if (r._GetLastNotEmptyColumn(true) > 0)
                    break;
                sheet.RemoveRow(r);
            }
        }
    }
}