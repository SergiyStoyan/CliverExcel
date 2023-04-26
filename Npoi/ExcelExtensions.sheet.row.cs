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
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
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
            for (int i = y1 - 1; i < y2; i++)
            {
                var r = sheet.GetRow(i);
                if (r == null)
                {
                    if (rowScope <= RowScope.ExistingOnly)
                        continue;
                    if (rowScope == RowScope.CreateIfNull)
                        r = sheet.CreateRow(i);
                }
                else if (r.LastCellNum < 0)
                {
                    if (rowScope <= RowScope.NotEmptyOnly)
                        continue;
                }
                else if (rowScope == RowScope.NotEmptyCellsOnly && r._GetLastNotEmptyColumn(false) < 1)
                    continue;
                yield return r;
            }
        }

        static public IRow _AppendRow<T>(this ISheet sheet, IEnumerable<T> values)
        {
            int lastRowY = sheet._GetLastNotEmptyRow(false);
            return sheet._WriteRow(lastRowY + 1, values);
        }

        static public IRow _AppendRow(this ISheet sheet, params string[] values)
        {
            return sheet._AppendRow(values);
        }

        static public IRow _InsertRow<T>(this ISheet sheet, int y, IEnumerable<T> values = null)
        {
            int lastRowY = sheet._GetLastNotEmptyRow(false);
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

        static public IRow _RemoveRow(this ISheet sheet, int y)
        {
            IRow r = sheet.GetRow(y - 1);
            if (r != null)
                sheet.RemoveRow(r);
            return r;
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
            sheet._GetRowsInRange(RowScope.ExistingOnly, y1, y2).ForEach(a => a.Height = -1);
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
            return sheet._GetRowsInRange(RowScope.ExistingOnly, y1, y2).Max(a => a._GetLastColumn(includeMerged));
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
            return sheet._GetRowsInRange(RowScope.ExistingOnly, y1, y2).Max(a => a._GetLastNotEmptyColumn(includeMerged));
        }
    }
}