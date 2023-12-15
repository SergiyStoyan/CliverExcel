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

            ///// <summary>         
            ///// NULL- and type-safe.
            ///// (!)Never returns NULL.
            ///// </summary>
            ///// <param name="y"></param>
            ///// <returns></returns>
            //public string this[int y]
            //{
            //    get
            //    {
            //        return GetValueAsString(y, false);
            //    }
            //    set
            //    {
            //        GetCell(y, true).SetCellValue(value);
            //    }
            //}

            public void SetAlteredStyles<T>(T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                var styleCache = Sheet.Workbook._Excel().OneWorkbookStyleCache;
                foreach (ICell cell in GetCells(true))
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
                var r = c._GetMergedRange();
                if (r != null)
                    return r.Y2.Value;
                return row._Y();
            }

            public IEnumerable<ICell> GetCells(bool createCells)
            {
                return GetCellsInRange(createCells);
            }

            public IEnumerable<ICell> GetCellsInRange(bool createCells, int y1 = 1, int? y2 = null)
            {
                if (y2 == null)
                    y2 = GetLastRow(LastRowCondition.HasCells, false);
                for (int y = y1; y <= y2; y++)
                    yield return GetCell(y, createCells);
            }

            public void SetStyles(int y1, IEnumerable<ICellStyle> styles)
            {
                SetStyles(y1, styles.ToArray());
            }

            public void SetStyles(int y1, params ICellStyle[] styles)
            {
                var cs = GetCellsInRange(true, y1, styles.Length).ToList();
                for (int i = y1 - 1; i < styles.Length; i++)
                    cs[i].CellStyle = styles[i];
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

            public IEnumerable<ICell> GetCells(RowScope rowScope)
            {
                return Sheet._GetRows(rowScope).Select(a => a?.GetCell(X));
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

            public int GetWidth()
            {
                return Sheet.GetColumnWidth(X - 1);
            }

            public void Copy(string column2Name, CopyCellMode copyCellMode = null)
            {
                Copy(Excel.GetX(column2Name), copyCellMode);
            }

            public void Copy(int x2, CopyCellMode copyCellMode = null)
            {
                if (X == x2)
                    return;
                Column column2 = new Column(Sheet, x2);
                column2.Clear(false);
                column2.SetWidth(GetWidth());
                foreach (ICell c1 in GetCells(false))
                    c1._Copy(x2, c1._X(), copyCellMode);
            }

            public void Move(string column2Name, MoveRegionMode moveRegionMode = null)
            {
                Move(Excel.GetX(column2Name), moveRegionMode);
            }

            /// <summary>
            /// Insert a copy and remove the source.
            /// </summary>
            /// <param name="x2"></param>
            /// <param name="moveRegionMode"></param>
            public void Move(int x2, MoveRegionMode moveRegionMode = null)
            {
                Sheet._ShiftColumnsRight(x2, 1, moveRegionMode);

                if (moveRegionMode?.UpdateMergedRegions == true)
                {
                    Sheet.MergedRegions.ForEach(a =>
                    {
                        if (a.FirstColumn < x2 - 1)
                        {
                            if (a.LastColumn >= x2 - 1)
                                a.LastColumn += 1;
                        }
                        else
                        {
                            a.FirstColumn += 1;
                            a.LastColumn += 1;
                        }
                    });
                }
                Copy(x2, moveRegionMode);
                Remove(moveRegionMode);
            }

            /// <summary>
            /// Remove the column from its sheet and shift columns on right. 
            /// </summary>
            /// <param name="moveRegionMode"></param>
            public void Remove(MoveRegionMode moveRegionMode = null)
            {
                if (moveRegionMode?.UpdateMergedRegions == true)
                {
                    for (int i = Sheet.MergedRegions.Count - 1; i >= 0; i--)
                    {
                        NPOI.SS.Util.CellRangeAddress a = Sheet.GetMergedRegion(i);
                        if (a.FirstColumn < X - 1)
                        {
                            if (a.LastColumn >= X - 1)
                                a.LastColumn -= 1;
                        }
                        else if (a.FirstColumn == X - 1 && a.LastColumn == X - 1)
                            Sheet.RemoveMergedRegion(i);
                        else
                        {
                            a.FirstColumn -= 1;
                            a.LastColumn -= 1;
                        }
                    }
                }
                Sheet._ShiftColumnsLeft(X + 1, 1, moveRegionMode);
            }

            public object GetValue(int y)
            {
                return GetCell(y, false)?._GetValue();
            }

            public string GetValueAsString(int y, bool allowNull = false)
            {
                ICell c = GetCell(y, false);
                if (c == null)
                    return allowNull ? null : string.Empty;
                return c._GetValueAsString(allowNull);
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