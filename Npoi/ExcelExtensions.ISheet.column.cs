﻿//********************************************************************************************
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
using System.Text.RegularExpressions;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public Column _AppendColumn<T>(this ISheet sheet, params T[] values)
        {
            return sheet._AppendColumn((IEnumerable<T>)values);
        }

        static public Column _AppendColumn<T>(this ISheet sheet, IEnumerable<T> values)
        {
            int x = sheet._GetLastColumn(false) + 1;
            Column c = new Column(sheet, x);
            c._Write(values);
            return c;
        }

        static public Column _InsertColumn<T>(this ISheet sheet, int x, IEnumerable<T> values = null, MoveRegionMode moveRegionMode = null)
        {
            sheet._ShiftColumnsRight(x, 1, moveRegionMode);
            Column c = new Column(sheet, x);
            c._Write(values);
            return c;
        }

        static public Column _InsertColumn<T>(this ISheet sheet, int x, params T[] values)
        {
            return sheet._InsertColumn(x, (IEnumerable<T>)values);
        }

        static public void _InsertColumnRange(this ISheet sheet, int x, int count, MoveRegionMode moveRegionMode = null)
        {
            sheet._ShiftColumnsRight(x, count, moveRegionMode);
        }

        static public Column _WriteColumn<T>(this ISheet sheet, int x, IEnumerable<T> values)
        {
            Column c = sheet._GetColumn(x);
            c._Write(values);
            return c;
        }

        static public Column _WriteColumn<T>(this ISheet sheet, int x, params T[] values)
        {
            return sheet._WriteColumn(x, (IEnumerable<T>)values);
        }

        static public void _RemoveColumn(this ISheet sheet, int x, MoveRegionMode moveRegionMode = null)
        {
            sheet._RemoveColumnRange(x, x, moveRegionMode);
        }

        static public void _RemoveColumnRange(this ISheet sheet, int x1, int x2, MoveRegionMode moveRegionMode = null)
        {
            sheet._ShiftColumnsLeft(x2 + 1, x2 - x1 + 1, moveRegionMode);
        }

        static public int _GetLastNotEmptyRowInColumnRange(this ISheet sheet, bool includeMerged, int x1 = 1, int? x2 = null)
        {
            if (x2 == null)
                x2 = int.MaxValue;
            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    continue;
                var c = row.Cells.Find(a => a.ColumnIndex + 1 >= x1 && a.ColumnIndex < x2 && !string.IsNullOrEmpty(a._GetValueAsString()));
                if (string.IsNullOrEmpty(c?._GetValueAsString()))
                    continue;
                if (includeMerged)
                {
                    var r = c._GetMergedRange();
                    if (r != null)
                        return r.Y2.Value;
                }
                return c.RowIndex + 1;
            }
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="xs"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyRowInColumns(this ISheet sheet, bool includeMerged, params int[] xs)
        {
            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    continue;
                var c = row.Cells.Find(a => xs.Contains(a.ColumnIndex + 1) && !string.IsNullOrEmpty(a._GetValueAsString()));
                if (string.IsNullOrEmpty(c?._GetValueAsString()))
                    continue;
                if (includeMerged)
                {
                    var r = c._GetMergedRange();
                    if (r != null)
                        return r.Y2.Value;
                }
                return c.RowIndex + 1;
            }
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="x"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastRowInColumn(this ISheet sheet, LastRowCondition lastRowCondition, bool includeMerged, int x)
        {
            return sheet._GetColumn(x).GetLastRow(lastRowCondition, includeMerged);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellValue"></param>
        /// <param name="cellY"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _FindColumnByCellValue(this ISheet sheet, Regex cellValueRegex, int cellY = 1)
        {
            IRow row = sheet._GetRow(cellY, false);
            if (row == null)
                return 0;
            for (int x = 1; x <= row.Cells.Count; x++)
                if (cellValueRegex.IsMatch(sheet._GetValueAsString(cellY, x, false)))
                    return x;
            return 0;
        }

        static public void _ShiftColumnsRight(this ISheet sheet, int x, int shift, MoveRegionMode moveRegionMode = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x;
            columnXs2width[lastColumnX] = sheet.GetColumnWidth(lastColumnX - 1);
            foreach (IRow row in sheet._GetRows(RowScope.NotNull))
            {
                int columnX = row._GetLastColumn(false);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = columnX; i >= x; i--)
                    sheet._MoveCell(row._Y(), i, row._Y(), i + shift, moveRegionMode);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                sheet._SetColumnWidth(columnX + shift, columnXs2width[columnX]);

            if (moveRegionMode?.UpdateMergedRegions == true)
                for (int i = sheet.MergedRegions.Count - 1; i >= 0; i--)
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(i);
                    if (a.FirstColumn < x - 1)
                    {
                        if (a.LastColumn < x - 1)
                        { }
                        else
                            a.LastColumn += shift;
                    }
                    else
                    {
                        a.FirstColumn += shift;
                        a.LastColumn += shift;
                    }
                }
        }

        static public void _ShiftColumnsLeft(this ISheet sheet, int x, int shift, MoveRegionMode moveRegionMode = null)
        {
            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
            int lastColumnX = x;
            columnXs2width[lastColumnX] = sheet.GetColumnWidth(lastColumnX - 1);
            foreach (IRow row in sheet._GetRows(RowScope.NotNull))
            {
                int columnX = row._GetLastColumn(false);
                if (lastColumnX < columnX)
                {
                    for (int i = lastColumnX; i < columnX; i++)
                        columnXs2width[i + 1] = sheet.GetColumnWidth(i);
                    lastColumnX = columnX;
                }
                for (int i = x; i <= columnX; i++)
                    sheet._MoveCell(row._Y(), i, row._Y(), i - shift, moveRegionMode);
            }
            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
                sheet._SetColumnWidth(columnX - shift, columnXs2width[columnX]);

            if (moveRegionMode?.UpdateMergedRegions == true)
                for (int i = sheet.MergedRegions.Count - 1; i >= 0; i--)
                {
                    NPOI.SS.Util.CellRangeAddress a = sheet.GetMergedRegion(i);
                    if (a.FirstColumn < x - 1)
                    {
                        if (a.LastColumn < x - 1)
                        { }
                        else if (a.LastColumn <= x - 1 + shift)
                            a.LastColumn = x - 1;
                        else
                            a.LastColumn -= shift;
                    }
                    else if (a.LastColumn <= x - 1 + shift)
                        sheet.RemoveMergedRegion(i);
                    else
                    {
                        a.FirstColumn -= shift;
                        a.LastColumn -= shift;
                    }
                }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyColumn(this ISheet sheet, bool includeMerged)
        {
            return sheet._GetLastNotEmptyColumnInRowRange(includeMerged, 1, null);
        }

        static public void _CopyColumn(this ISheet sheet, string Column1Name, string Column2Name, CopyCellMode copyCellMode = null)
        {
            sheet._GetColumn(Column1Name).Copy(Column2Name, copyCellMode);
        }

        static public void _CopyColumn(this ISheet sheet, int x1, int x2, CopyCellMode copyCellMode = null)
        {
            sheet._GetColumn(x1).Copy(x2, copyCellMode);
        }

        static public void _MoveColumn(this ISheet sheet, string Column1Name, string Column2Name, bool insert, MoveRegionMode moveRegionMode = null)
        {
            sheet._GetColumn(Column1Name).Move(Column2Name, insert, moveRegionMode);
        }

        static public void _MoveColumn(this ISheet sheet, int x1, int x2, bool insert, MoveRegionMode moveRegionMode = null)
        {
            var c1 = sheet._GetColumn(x1);
            c1.Move(x2, insert, moveRegionMode);
        }

        static public int _GetLastNotEmptyRowInColumn(this ISheet sheet, bool includeMerged, int x)
        {
            return sheet._GetColumn(x).GetLastNotEmptyRow(includeMerged);
        }

        static public Column _GetColumn(this ISheet sheet, int x)
        {
            return new Column(sheet, x);
        }

        static public Column _GetColumn(this ISheet sheet, string columnName)
        {
            return new Column(sheet, CellReference.ConvertColStringToIndex(columnName));
        }

        static public IEnumerable<Column> _GetColumns(this ISheet sheet)
        {
            return sheet._GetColumnsInRange();
        }

        static public IEnumerable<Column> _GetColumnsInRange(this ISheet sheet, int x1 = 1, int? x2 = null)
        {
            if (x2 == null)
                x2 = sheet._GetLastColumn(false);
            for (int x = x1; x <= x2; x++)
                yield return new Column(sheet, x);
        }

        static public int _GetLastColumn(this ISheet sheet, bool includeMerged)
        {
            return sheet._GetLastColumnInRowRange(includeMerged, 1, null);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int _GetLastNotEmptyRow(this ISheet sheet, bool includeMerged)
        {
            return sheet._GetLastNotEmptyRowInColumnRange(includeMerged, 1, null);
        }

        static public void _ShiftColumnCellsDown(this ISheet sheet, int x, int y1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetColumn(x)?.ShiftCellsDown(y1, shift, copyCellMode);
        }

        static public void _ShiftColumnCellsUp(this ISheet sheet, int x, int y1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetColumn(x)?.ShiftCellsUp(y1, shift, copyCellMode);
        }

        static public void _ShiftColumnCells(this ISheet sheet, int x, int y1, int shift, CopyCellMode copyCellMode = null)
        {
            sheet._GetColumn(x)?.ShiftCells(y1, shift, copyCellMode);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="padding">a character width</param>
        static public void _AutosizeColumns(this ISheet sheet, float padding = 0)
        {
            sheet._AutosizeColumnsInRange(1, null, padding);
        }

        static public void _ClearColumnRange(this ISheet sheet, int x1, int x2, bool clearMerging)
        {
            for (int x = x1; x <= x2; x++)
                sheet._ClearColumn(x, clearMerging);
        }

        static public void _ClearColumn(this ISheet sheet, int x, bool clearMerging)
        {
            sheet._GetColumn(x).Clear(clearMerging);
        }

        static public void _ClearMergingInColumn(this ISheet sheet, int x)
        {
            sheet._NewRange(1, x, null, x).ClearMerging();
        }

        static public void _SetStyleInColumn(this ISheet sheet, ICellStyle style, bool createCells, int x)
        {
            sheet._SetStyleInColumnRange(style, createCells, x, x);
        }

        static public void _SetStyleInColumnRange(this ISheet sheet, ICellStyle style, bool createCells, int x1, int? x2 = null)
        {
            sheet._NewRange(1, x1, null, x2).SetStyle(style, createCells);
        }

        static public void _ReplaceStyleInColumnRange(this ISheet sheet, ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
        {
            sheet._NewRange(1, x1, null, x2).ReplaceStyle(style1, style2);
        }

        static public void _ClearStyleInColumnRange(this ISheet sheet, ICellStyle style, int x1, int? x2 = null)
        {
            sheet._ReplaceStyleInColumnRange(style, null, x1, x2);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="x2"></param>
        /// <param name="padding">a character width</param>
        static public void _AutosizeColumnsInRange(this ISheet sheet, int x1 = 1, int? x2 = null, float padding = 0)
        {
            if (x2 == null)
                x2 = sheet._GetLastColumn(false);
            for (int x = x1; x <= x2; x++)
                sheet._AutosizeColumn(x, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="columnIs"></param>
        /// <param name="padding">a character width</param>
        static public void _AutosizeColumns(this ISheet sheet, IEnumerable<int> Xs, float padding = 0)
        {
            foreach (int y in Xs)
                sheet._AutosizeColumn(y, padding);
        }

        /// <summary>
        /// (!)Very slow on large data.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="padding">a character width</param>
        static public void _AutosizeColumn(this ISheet sheet, int x, float padding = 0)
        {
            sheet._GetColumn(x).Autosize(padding);
        }

        static public IEnumerable<ICell> _GetCellsInColumn(this ISheet sheet, int x, RowScope rowScope)
        {
            return sheet._GetColumn(x).GetCells(rowScope);
        }

        /// <summary>
        /// Safe comparing to the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">units of 1/256th of a character width</param>
        static public void _SetColumnWidth(this ISheet sheet, int x, int width)
        {
            sheet._GetColumn(x).SetWidth(width);
        }

        /// <summary>
        /// Safe comparing to the API's one
        /// </summary>
        /// <param name="x"></param>
        /// <param name="width">a character width</param>
        static public void _SetColumnWidth(this ISheet sheet, int x, float width)
        {
            sheet._GetColumn(x).SetWidth(width);
        }
    }
}