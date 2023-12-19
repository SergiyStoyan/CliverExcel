//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel
    {
        public class Column
        {
            internal protected Column(ISheet sheet, int x)
            {
                Sheet = sheet;
                X = x;
            }
            readonly public ISheet Sheet;

            public readonly int X;

            public ICell GetCell(int y, bool createCell)
            {
                return Sheet._GetCell(y, X, createCell);
            }

            public void SetAlteredStyles<T>(T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                var styleCache = Sheet.Workbook._Excel().OneWorkbookStyleCache;
                foreach (ICell cell in GetCells(CellScope.CreateIfNull))
                    cell.CellStyle = styleCache.GetAlteredStyle(cell.CellStyle, alterationKey, alterStyle, reuseUnusedStyle);
            }

            public int GetLastRow(LastRowCondition lastRowCondition, bool includeMerged)
            {
                IRow row = null;
                switch (lastRowCondition)
                {
                    case LastRowCondition.NotEmpty:
                        return GetLastNotEmptyRow(includeMerged);
                    case LastRowCondition.HasCells:
                        for (int i = Sheet.LastRowNum; i >= 0; i--)
                        {
                            row = Sheet.GetRow(i);
                            if (row == null)
                                continue;
                            if (row.GetCell(X - 1) != null)
                                break;
                        }
                        break;
                    case LastRowCondition.NotNull:
                        row = Sheet.GetRow(Sheet.LastRowNum);
                        break;
                    default:
                        throw new Exception("Unknown option: " + lastRowCondition.ToString());
                }
                if (row == null)
                    return 0;
                if (!includeMerged)
                    return row._Y();
                var c = row.GetCell(X - 1);
                var r = c?._GetMergedRange();
                if (r != null)
                    return r.Y2.Value;
                return row._Y();
            }

            public IEnumerable<ICell> GetCells(CellScope cellScope)
            {
                return GetCellsInRange(cellScope);
            }

            public IEnumerable<ICell> GetCellsInRange(CellScope cellScope, int y1 = 1, int? y2 = null)
            {
                if (y2 == null)
                    y2 = Sheet.LastRowNum + 1; //GetLastRow(LastRowCondition.HasCells, false);
                switch (cellScope)
                {
                    case CellScope.NotEmpty:
                        for (int y = y1; y <= y2; y++)
                        {
                            var c = GetCell(y, false);
                            if (!string.IsNullOrWhiteSpace(c._GetValueAsString()))
                                yield return c;
                        }
                        break;
                    case CellScope.NotNull:
                        for (int y = y1; y <= y2; y++)
                        {
                            var c = GetCell(y, false);
                            if (c != null)
                                yield return c;
                        }
                        break;
                    case CellScope.IncludeNull:
                        for (int y = y1; y <= y2; y++)
                        {
                            var c = GetCell(y, false);
                            yield return c;
                        }
                        break;
                    case CellScope.CreateIfNull:
                        for (int y = y1; y <= y2; y++)
                        {
                            var c = GetCell(y, true);
                            yield return c;
                        }
                        break;
                    default: throw new Exception("Unknown option: " + cellScope.ToString());
                }
            }

            public void SetStyles(int y1, IEnumerable<ICellStyle> styles)
            {
                SetStyles(y1, styles.ToArray());
            }

            public void SetStyles(int y1, params ICellStyle[] styles)
            {
                for (int i = y1 - 1; i < styles.Length; i++)
                    GetCell(i + 1, true).CellStyle = styles[i];
            }

            public void SetStyle(ICellStyle style, bool createCells)
            {
                new Range(Sheet, 1, X, null, X).SetStyle(style, createCells);
            }

            public int GetLastNotEmptyRow(bool includeMerged)
            {
                for (int i = Sheet.LastRowNum; i >= 0; i--)
                {
                    IRow row = Sheet.GetRow(i);
                    if (row == null)
                        continue;
                    var c = row.GetCell(X - 1);
                    if (string.IsNullOrEmpty(c._GetValueAsString()))
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

            public void Clear(bool clearMerging)
            {
                if (clearMerging)
                    ClearMerging();
                foreach (var r in Sheet._GetRows(RowScope.NotNull))
                {
                    var c = r.GetCell(X - 1);
                    if (c != null)
                        r.RemoveCell(c);
                }
            }

            public void ClearMerging()
            {
                new Range(Sheet, 1, X, null, X).ClearMerging();
            }

            public void ShiftCellsDown(int y1, int shift, CopyCellMode copyCellMode = null)
            {
                if (shift < 0)
                    throw new Exception("Shift cannot be < 0: " + shift);
                for (int y = GetLastRow(LastRowCondition.HasCells, true); y >= y1; y--)
                    Sheet._MoveCell(y, X, y + shift, X, copyCellMode);
            }

            public void ShiftCellsUp(int y1, int shift, CopyCellMode copyCellMode = null)
            {
                if (shift < 0)
                    throw new Exception("Shift cannot be < 0: " + shift);
                if (shift >= y1)
                    throw new Exception("Shifting up before the first row: shift=" + shift + ", y1=" + y1);
                int y2 = GetLastRow(LastRowCondition.HasCells, true) + 1;
                for (int y = y1; y <= y2; y++)
                    Sheet._MoveCell(y, X, y - shift, X, copyCellMode);
            }

            public void ShiftCells(int y1, int shift, CopyCellMode copyCellMode = null)
            {
                if (shift >= 0)
                    ShiftCellsUp(y1, shift, copyCellMode);
                else
                    ShiftCellsDown(y1, -shift, copyCellMode);
            }

            /// <summary>
            /// Safe comparing to the API's one
            /// </summary>
            /// <param name="x"></param>
            /// <param name="width">units of 1/256th of a character width</param>
            public void SetWidth(int width)
            {
                const int cellMaxWidth = 256 * 255;
                int w = MathRoutines.Truncate(width, cellMaxWidth);
                Sheet.SetColumnWidth(X - 1, w);
            }

            /// <summary>
            /// Safe comparing to the API's one
            /// </summary>
            /// <param name="x"></param>
            /// <param name="width">a character width</param>
            public void SetWidth(float width)
            {
                SetWidth((int)(width * 256));
            }

            public int GetWidth()
            {
                return Sheet.GetColumnWidth(X - 1);
            }

            /// <summary>
            /// (!)Very slow on large data.
            /// </summary>
            /// <param name="x"></param>
            /// <param name="padding">a character width</param>
            public void Autosize(float padding = 0)
            {
                Sheet.AutoSizeColumn(X - 1, false);

                //GetCellsInColumn(x).Max(a => a.GetValueAsString())
                //int width = ((int)(maxNumCharacters * 1.14388)) * 256;
                //sheet.setColumnWidth(i, width);

                if (padding > 0)
                    SetWidth(Sheet.GetColumnWidth(X - 1) + (int)(padding * 256));
            }

            public Column Copy(string column2Name, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._CopyColumn(X, GetX(column2Name), copyCellMode, sheet2, styleMap);
            }

            public Column Copy(int x2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._CopyColumn(X, x2, copyCellMode, sheet2, styleMap);
            }

            public Column Copy(Column column2, CopyCellMode copyCellMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._CopyColumn(X, column2.X, copyCellMode, sheet2, styleMap);
            }

            public Column Move(string column2Name, bool insert, MoveRegionMode moveRegionMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._MoveColumn(X, GetX(column2Name), insert, moveRegionMode, sheet2, styleMap);
            }

            public Column Move(Column column2, bool insert, MoveRegionMode moveRegionMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._MoveColumn(X, column2.X, insert, moveRegionMode, sheet2, styleMap);
            }

            /// <summary>
            /// Insert a copy and remove the source.
            /// </summary>
            /// <param name="x2"></param>
            /// <param name="moveRegionMode"></param>
            public Column Move(int x2, bool insert, MoveRegionMode moveRegionMode = null, ISheet sheet2 = null, StyleMap styleMap = null)
            {
                return Sheet._MoveColumn(X, x2, insert, moveRegionMode, sheet2, styleMap);
            }

            /// <summary>
            /// Remove the column from its sheet and shift columns on right. 
            /// </summary>
            /// <param name="moveRegionMode"></param>
            public void Remove(MoveRegionMode moveRegionMode = null)
            {
                Sheet._RemoveColumn(X, moveRegionMode);
            }

            public object GetValue(int y)
            {
                return GetCell(y, false)?._GetValue();
            }

            public string GetValueAsString(int y, StringMode stringMode = DefaultStringMode)
            {
                ICell c = GetCell(y, false);
                return c._GetValueAsString(stringMode);
            }

            /// <summary>
            /// Images anchored in the specified cell coordinates. The cell may possibly not exist.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="x"></param>
            /// <returns></returns>
            public IEnumerable<Excel.Image> GetImages(int y)
            {
                return Sheet._GetImages(y, X);
            }

            public void _Write<T>(IEnumerable<T> values)
            {
                int y = 1;
                foreach (T v in values)
                {
                    if (v != null)
                        Sheet._GetCell(y, X, true)._SetValue(v);
                    y++;
                }
            }

            public void _Write(params string[] values)
            {
                _Write((IEnumerable<string>)values);
            }
        }
    }
}