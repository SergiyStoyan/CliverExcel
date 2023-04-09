<<<<<<< Updated upstream
﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        //static public ICell GetCell(this IRow r, string header, bool create)
        //{
        //    ICell c = r.GetCell(x - 1);
        //    if (c == null && create)
        //        return r.CreateCell(x - 1);
        //    return c;
        //}

        static public ICell GetCell(this IRow r, int x, bool create)
        {
            ICell c = r.GetCell(x - 1);
            if (c == null && create)
                return r.CreateCell(x - 1);
            return c;
        }
=======
﻿////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************
//using NPOI.SS.UserModel;
//using System;
//using System.Collections.Generic;
//using System.Linq;

//namespace Cliver
//{
//    static public partial class ExcelExtensions
//    {
//        static public void ShiftCellsRight(, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            for (int x = row.GetLastColumn(true); x >= x1; x--)
//                row.Sheet.MoveCell(row.Y(), x, row.Y(), x + shift, onFormulaCellMoved);
//        }

//        static public void ShiftCellsLeft(, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            int x2 = row.GetLastColumn(true);
//            for (int x = x1; x <= x2; x++)
//                row.Sheet.MoveCell(row.Y(), x, row.Y(), x - shift, onFormulaCellMoved);
//        }

//        //static public ICell GetCell(this IRow r, string header, bool create)
//        //{
//        //    ICell c = r.GetCell(x - 1);
//        //    if (c == null && create)
//        //        return r.CreateCell(x - 1);
//        //    return c;
//        //}

//        static public ICell GetCell(this IRow r, int x, bool createCell)
//        {
//            ICell c = r.GetCell(x - 1);
//            if (c == null && createCell)
//                return r.CreateCell(x - 1);
//            return c;
//        }
>>>>>>> Stashed changes

//        //static public void Highlight(, ICellStyle style, Excel.Color color)
//        //{
//        //    row.RowStyle = Excel.highlight(row.Sheet.Workbook, style, color);
//        //}

<<<<<<< Updated upstream
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="includeMerged"></param>
        /// <returns>1-based, otherwise 0</returns>
        static public int GetLastNotEmptyColumn(this IRow row, bool includeMerged = true)
        {
            if (row == null || row.Cells.Count < 1)
                return 0;
            for (int x0 = row.Cells.Count - 1; x0 >= 0; x0--)
            {
                var c = row.Cells[x0];
                if (!string.IsNullOrWhiteSpace(c.GetValueAsString()))
                {
                    if (includeMerged)
                    {
                        var r = c.GetMergedRange();
                        if (r != null)
                            return r.X2.Value;
                    }
                    return c.ColumnIndex + 1;
                }
            }
            return 0;
        }
=======
//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="row"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int GetLastNotEmptyColumn(, bool includeMerged = true)
//        {
//            if (row == null || row.Cells.Count < 1)
//                return 0;
//            for (int x0 = row.Cells.Count - 1; x0 >= 0; x0--)
//            {
//                var c = row.GetCell(x0);
//                if (!string.IsNullOrWhiteSpace(c?.GetValueAsString()))
//                {
//                    if (includeMerged)
//                    {
//                        var r = c.GetMergedRange();
//                        if (r != null)
//                            return r.X2.Value;
//                    }
//                    return c.ColumnIndex + 1;
//                }
//            }
//            return 0;
//        }
>>>>>>> Stashed changes

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="row"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int GetLastColumn(, bool includeMerged = true)
//        {
//            if (row == null || row.Cells.Count < 1)
//                return 0;
//            if (includeMerged)
//            {
//                var c = row.Cells[row.Cells.Count - 1];
//                var r = c.GetMergedRange();
//                if (r != null)
//                    return r.X2.Value;
//                return c.ColumnIndex + 1;
//            }
//            return row.LastCellNum;
//        }

//        static public IEnumerable<ICell> GetCells(, bool createCells)
//        {
//            return GetCellsInRange(row, createCells);
//        }

//        static public IEnumerable<ICell> GetCellsInRange(, bool createCells, int x1 = 1, int? x2 = null)
//        {
//            if (row == null)
//                yield break;
//            if (x2 == null)
//                x2 = row.LastCellNum;
//            for (int x = x1; x <= x2; x++)
//                yield return row.GetCell(x, createCells);
//        }

//        /// <summary>
//        /// 1-based row index on the sheet.
//        /// </summary>
//        /// <param name="row"></param>
//        /// <returns>1-based</returns>
//        static public int Y()
//        {
//            return row.RowNum + 1;
//        }

//        static public void Write(, IEnumerable<object> values)
//        {
//            int x = 1;
//            foreach (object v in values)
//                row.GetCell(x++, true).SetValue(v);
//        }

//        static public void Write(, params object[] values)
//        {
//            Write(row, (IEnumerable<object>)values);
//        }

//        static public void SetStyles(, int x1, IEnumerable<ICellStyle> styles)
//        {
//            SetStyles(row, x1, styles.ToArray());
//        }

<<<<<<< Updated upstream
        static public void SetStyles(this IRow row, int x1, params ICellStyle[] styles)
        {
            var cs = row.GetCellsInRange(true, x1, styles.Length).ToList();
            for (int i = x1 - 1; i < styles.Length; i++)
                cs[i].CellStyle = styles[i];
        }
    }
}
=======
//        static public void SetStyles(, int x1, params ICellStyle[] styles)
//        {
//            var cs = row.GetCellsInRange(true, x1, styles.Length).ToList();
//            for (int i = x1 - 1; i < styles.Length; i++)
//                cs[i].CellStyle = styles[i];
//        }


//        static public void Clear(, int y, bool clearMerging)
//        {
//            if (clearMerging)
//                row.ClearMerging();
//            row.Sheet.RemoveRow(row);
//        }

//        static public void ClearMerging()
//        {
//            new Excel.Range(row.Sheet, row.Y(), 1, row.Y(), null).ClearMerging();
//        }
//    }
//}
>>>>>>> Stashed changes
