////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************

//using System;
//using System.Collections.Generic;
//using NPOI.SS.UserModel;
//using static Cliver.Excel;
//using System.Linq;
//using NPOI.SS.Util;
//using NPOI.XSSF.UserModel;

//namespace Cliver
//{
//    static public partial class ExcelExtensions
//    {
//        static public IRow GetRow(this ISheet sheet, int y, bool createRow)
//        {
//            IRow r = sheet.GetRow(y - 1);
//            if (r == null && createRow)
//                r = sheet.CreateRow(y - 1);
//            return r;
//        }

//        static public IEnumerable<IRow> GetRows(this ISheet sheet, RowScope rowScope = RowScope.IncludeNull)
//        {
//            return sheet.GetRowsInRange(rowScope);
//        }

//        static public IEnumerable<IRow> GetRowsInRange(this ISheet sheet, RowScope rowScope = RowScope.IncludeNull, int y1 = 1, int? y2 = null)
//        {
//            if (y2 == null)
//                y2 = sheet.LastRowNum + 1;
//            //var rows = Sheet.GetRowEnumerator();//!!!buggy: sometimes misses added rows
//            for (int i = y1 - 1; i < y2; i++)
//            {
//                var r = sheet.GetRow(i);
//                if (r == null)
//                {
//                    if (rowScope == RowScope.OnlyExisting)
//                        continue;
//                    if (rowScope == RowScope.CreateIfNull)
//                        r = sheet.CreateRow(i);
//                }
//                if (r != null)
//                    yield return r;
//            }
//        }

//        static public IRow AppendRow<T>(this ISheet sheet, IEnumerable<T> values)
//        {
//            int y0 = sheet.LastRowNum;//(!)it is 0 when no row or 1 row
//            int y = y0 + (y0 == 0 && sheet.GetRow(y0) == null ? 1 : 2);
//            return sheet.WriteRow(y, values);
//        }

//        static public IRow AppendRow<T>(this ISheet sheet, params T[] values)
//        {
//            return sheet.AppendRow(values);
//        }

//        static public IRow InsertRow<T>(this ISheet sheet, int y, IEnumerable<T> values = null)
//        {
//            if (y <= sheet.LastRowNum)
//                sheet.ShiftRows(y - 1, sheet.LastRowNum, 1);
//            return sheet.WriteRow(y, values);
//        }

//        static public IRow InsertRow<T>(this ISheet sheet, int y, params T[] values)
//        {
//            return sheet.InsertRow(y, (IEnumerable<T>)values);
//        }

//        static public IRow WriteRow<T>(this ISheet sheet, int y, IEnumerable<T> values)
//        {
//            IRow r = sheet.GetRow(y, true);
//            r.Write(values);
//            return r;
//        }

//        static public IRow WriteRow<T>(this ISheet sheet, int y, params T[] values)
//        {
//            return sheet.WriteRow(y, (IEnumerable<T>)values);
//        }

//        static public void ShiftRowCellsRight(this ISheet sheet, int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            sheet.GetRow(y, false)?.ShiftCellsRight(x1, shift, onFormulaCellMoved);
//        }

//        static public void ShiftRowCellsLeft(this ISheet sheet, int y, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            sheet.GetRow(y, false)?.ShiftCellsLeft(x1, shift, onFormulaCellMoved);
//        }

//        static public void SetStyleInRow(this ISheet sheet, ICellStyle style, bool createCells, int y)
//        {
//            sheet.SetStyleInRowRange(style, createCells, y, y);
//        }

//        static public void SetStyleInRowRange(this ISheet sheet, ICellStyle style, bool createCells, int y1, int? y2 = null)
//        {
//            sheet.NewRange(y1, 1, y2, null).SetStyle(style, createCells);
//        }

//        static public void ReplaceStyleInRowRange(this ISheet sheet, ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
//        {
//            sheet.NewRange(y1, 1, y2, null).ReplaceStyle(style1, style2);
//        }

//        static public void ClearStyleInRowRange(this ISheet sheet, ICellStyle style, int y1, int? y2 = null)
//        {
//            sheet.ReplaceStyleInRowRange(style, null, y1, y2);
//        }

//        static public void AutosizeRowsInRange(this ISheet sheet, int y1 = 1, int? y2 = null)
//        {
//            sheet.GetRowsInRange(RowScope.OnlyExisting, y1, y2).ForEach(a => a.Height = -1);
//        }

//        static public void AutosizeRows(this ISheet sheet)
//        {
//            sheet.AutosizeRowsInRange();
//        }

//        static public void ClearRow(this ISheet sheet, int y, bool clearMerging)
//        {
//            if (clearMerging)
//                sheet.ClearMergingInRow(y);
//            var r = sheet.GetRow(y, false);
//            if (r != null)
//                sheet.RemoveRow(r);
//        }

//        static public void ClearMergingInRow(this ISheet sheet, int y)
//        {
//            sheet.NewRange(y, 1, y, null).ClearMerging();
//        }

//        static public int GetLastRow(this ISheet sheet, bool includeMerged = true)
//        {
//            IRow row = sheet.GetRow(sheet.LastRowNum);
//            if (row == null)
//                return 0;
//            if (!includeMerged)
//                return row.Y();
//            int maxY = 0;
//            foreach (var c in row.Cells)
//            {
//                var r = c.GetMergedRange();
//                if (r != null && maxY < r.Y2.Value)
//                    maxY = r.Y2.Value;
//            }
//            return maxY;
//        }

//        static public int GetLastColumnInRowRange(this ISheet sheet, int y1 = 1, int? y2 = null, bool includeMerged = true)
//        {
//            return sheet.GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastColumn(includeMerged));
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="y"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int GetLastNotEmptyColumnInRow(this ISheet sheet, int y, bool includeMerged = true)
//        {
//            IRow row = sheet.GetRow(y, false);
//            if (row == null)
//                return 0;
//            return row.GetLastNotEmptyColumn(includeMerged);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="y"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int GetLastColumnInRow(this ISheet sheet, int y, bool includeMerged = true)
//        {
//            IRow row = sheet.GetRow(y, false);
//            if (row == null)
//                return 0;
//            return row.GetLastColumn(includeMerged);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="y1"></param>
//        /// <param name="y2"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int GetLastNotEmptyColumnInRowRange(this ISheet sheet, int y1 = 1, int? y2 = null, bool includeMerged = true)
//        {
//            if (y2 == null)
//                y2 = sheet.LastRowNum + 1;
//            return sheet.GetRowsInRange(RowScope.OnlyExisting, y1, y2).Max(a => a.GetLastNotEmptyColumn(includeMerged));
//        }
//    }
//}