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

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        /// <summary>
        /// Removes empty rows.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="includeEmptyCellRows">might considerably slow down if TRUE</param>
        /// <param name="shiftRemainingRows">the remaining rows are shifted if TRUE</param>
        public static void _RemoveEmptyRows(this ISheet sheet, bool includeEmptyCellRows, bool shiftRemainingRows)
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
                if (shiftRemainingRows)
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

        static public IRow _AppendRow<T>(this ISheet sheet, IEnumerable<T> values)
        {
            int lastRowY = sheet._GetLastRow(LastRowCondition.HasCells, false);
            return sheet._WriteRow(lastRowY + 1, values);
        }

        static public IRow _AppendRow(this ISheet sheet, params string[] values)
        {
            return sheet._AppendRow(values);
        }

        static public IRow _InsertRow<T>(this ISheet sheet, int y, IEnumerable<T> values = null)
        {
            int lastRowY = sheet._GetLastRow(LastRowCondition.HasCells, false);
            if (y <= lastRowY)
                sheet.ShiftRows(y - 1, lastRowY - 1, 1);
            return sheet._WriteRow(y, values);
        }

        static public IRow _InsertRow(this ISheet sheet, int y, params string[] values)
        {
            return sheet._InsertRow(y, (IEnumerable<string>)values);
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

        static public IRow _RemoveRow(this ISheet sheet, int y, bool shiftRemainingRows)
        {
            IRow r = sheet.GetRow(y - 1);
            if (r != null)
                sheet.RemoveRow(r);
            if (shiftRemainingRows)
                sheet.ShiftRows(r.RowNum + 1, sheet.LastRowNum, -1);
            return r;
        }

        static public void _MoveRow(this ISheet sheet, int y1, int y2)
        {
            if (y1 == y2)
                return;
            sheet.ShiftRows(y2 - 1, sheet.LastRowNum, 1);
            if (y1 > y2)
            {
                sheet.ShiftRows(y1, y1, y2 - y1 - 1);
                sheet.ShiftRows(y1 + 1, sheet.LastRowNum, -1);
            }
            else
            {
                sheet.ShiftRows(y1 - 1, y1 - 1, y2 - y1);
                sheet.ShiftRows(y1, sheet.LastRowNum, -1);
            }
        }

        static public void _ShiftRowCellsRight(this ISheet sheet, int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            sheet._GetRow(y, false)?._ShiftCellsRight(x1, shift, onFormulaCellMoved);
        }

        static public void _ShiftRowCellsLeft(this ISheet sheet, int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            sheet._GetRow(y, false)?._ShiftCellsLeft(x1, shift, onFormulaCellMoved);
        }

        static public void _SetStyleInRow(this ISheet sheet, ICellStyle style, bool createCells, int y)
        {
            sheet._SetStyleInRowRange(style, createCells, y, y);
        }

        static public void _SetStyleInRowRange(this ISheet sheet, ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            sheet._NewRange(y1, 1, y2, null).SetStyle(style, createCells);
        }

        static public void _ReplaceStyleInRowRange(this ISheet sheet, ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            sheet._NewRange(y1, 1, y2, null).ReplaceStyle(style1, style2);
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

        static public void _ClearRow(this ISheet sheet, int y, bool clearMerging)
        {
            if (clearMerging)
                sheet._ClearMergingInRow(y);
            var r = sheet._GetRow(y, false);
            if (r != null)
                sheet.RemoveRow(r);
        }

        static public void _ClearMergingInRow(this ISheet sheet, int y)
        {
            sheet._NewRange(y, 1, y, null).ClearMerging();
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
    }
}