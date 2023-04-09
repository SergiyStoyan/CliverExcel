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
//using System.Text.RegularExpressions;

//namespace Cliver
//{
//    static public partial class ExcelExtensions
//    {
//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="x"></param>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int _GetLastRowInColumn(this ISheet sheet, int x, bool includeMerged = true)
//        {
//            return sheet._GetColumn(x).GetLastRow(includeMerged);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="cellValue"></param>
//        /// <param name="cellY"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int _FindColumnByCellValue(this ISheet sheet, Regex cellValue, int cellY = 1)
//        {
//            IRow row = sheet.GetRow(cellY, false);
//            if (row == null)
//                return 0;
//            for (int x = 1; x <= row.Cells.Count; x++)
//                if (cellValue.IsMatch(sheet.GetValueAsString(cellY, x, false)))
//                    return x;
//            return 0;
//        }

//        static public void _ShiftColumnsRight(this ISheet sheet, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
//            int lastColumnX = x1;
//            columnXs2width[lastColumnX] = sheet.GetColumnWidth(lastColumnX - 1);
//            //var rows = Sheet._GetRowEnumerator();//!!!buggy: sometimes misses added rows
//            //while (rows.MoveNext())
//            for (int y0 = sheet.LastRowNum; y0 >= 0; y0--)
//            {
//                IRow row = sheet.GetRow(y0);
//                if (row == null)
//                    continue;
//                int columnX = row.GetLastColumn(true);
//                if (lastColumnX < columnX)
//                {
//                    for (int i = lastColumnX; i < columnX; i++)
//                        columnXs2width[i + 1] = sheet.GetColumnWidth(i);
//                    lastColumnX = columnX;
//                }
//                for (int i = columnX; i >= x1; i--)
//                    sheet._MoveCell(row.Y(), i, row.Y(), i + shift, onFormulaCellMoved);
//            }
//            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
//                sheet._SetColumnWidth(columnX + shift, columnXs2width[columnX]);
//        }

//        static public void _ShiftColumnsLeft(this ISheet sheet, int x1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            Dictionary<int, int> columnXs2width = new Dictionary<int, int>();
//            int lastColumnX = x1;
//            columnXs2width[lastColumnX] = sheet.GetColumnWidth(lastColumnX - 1);
//            //var rows = Sheet._GetRowEnumerator();//!!!buggy: sometimes misses added rows
//            //while (rows.MoveNext())
//            for (int y0 = sheet.LastRowNum; y0 >= 0; y0--)
//            {
//                IRow row = sheet.GetRow(y0);
//                if (row == null)
//                    continue;
//                int columnX = row.GetLastColumn(true);
//                if (lastColumnX < columnX)
//                {
//                    for (int i = lastColumnX; i < columnX; i++)
//                        columnXs2width[i + 1] = sheet.GetColumnWidth(i);
//                    lastColumnX = columnX;
//                }
//                for (int i = x1; i <= columnX; i++)
//                    sheet._MoveCell(row.Y(), i, row.Y(), i - shift, onFormulaCellMoved);
//            }
//            foreach (int columnX in columnXs2width.Keys.OrderByDescending(a => a))
//                sheet._SetColumnWidth(columnX - shift, columnXs2width[columnX]);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int _GetLastNotEmptyColumn(this ISheet sheet, bool includeMerged)
//        {
//            return sheet._GetLastNotEmptyColumnInRowRange(1, null, includeMerged);
//        }

//        static public void _CopyColumn(this ISheet sheet, string fromColumnName, ISheet toSheet, string toColumnName = null)
//        {
//            sheet._GetColumn(fromColumnName).Copy(toSheet, toColumnName);
//        }

//        static public void _CopyColumn(this ISheet sheet, int fromX, ISheet toSheet, int toX)
//        {
//            sheet._GetColumn(fromX).Copy(toSheet, toX);
//        }

//        static public int _GetLastNotEmptyRowInColumn(this ISheet sheet, int x, bool includeMerged = true)
//        {
//            return sheet._GetColumn(x).GetLastNotEmptyRow(includeMerged);
//        }
//        static public Column _GetColumn(this ISheet sheet, int x)
//        {
//            return new Column(sheet, x);
//        }

//        static public Column _GetColumn(this ISheet sheet, string columnName)
//        {
//            return new Column(sheet, CellReference.ConvertColStringToIndex(columnName));
//        }

//        static public IEnumerable<Column> _GetColumns(this ISheet sheet)
//        {
//            return sheet._GetColumnsInRange();
//        }

//        static public IEnumerable<Column> _GetColumnsInRange(this ISheet sheet, int x1 = 1, int? x2 = null)
//        {
//            if (x2 == null)
//                x2 = sheet._GetLastColumn(false);
//            for (int x = x1; x <= x2; x++)
//                yield return new Column(sheet, x);
//        }

//        static public int _GetLastColumn(this ISheet sheet, bool includeMerged = true)
//        {
//            return sheet._GetLastColumnInRowRange(1, null, includeMerged);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="includeMerged"></param>
//        /// <returns>1-based, otherwise 0</returns>
//        static public int _GetLastNotEmptyRow(this ISheet sheet, bool includeMerged = true)
//        {
//            return sheet._GetLastNotEmptyRowInColumnRange(1, null, includeMerged);
//        }

//        static public void _ShiftColumnCellsDown(this ISheet sheet, int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            sheet._GetColumn(x).ShiftCellsDown(y1, shift, onFormulaCellMoved);
//        }

//        static public void _ShiftColumnCellsUp(this ISheet sheet, int x, int y1, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            sheet._GetColumn(x).ShiftCellsUp(y1, shift, onFormulaCellMoved);
//        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <param name="padding">a character width</param>
//        static public void _AutosizeColumns(this ISheet sheet, float padding = 0)
//        {
//            sheet._AutosizeColumnsInRange(1, null, padding);
//        }

//        static public void _ClearColumn(this ISheet sheet, int x, bool clearMerging)
//        {
//            sheet._GetColumn(x).Clear(clearMerging);
//        }

//        static public void _ClearMergingInColumn(this ISheet sheet, int x)
//        {
//            sheet.NewRange(1, x, null, x).ClearMerging();
//        }

//        static public void _SetStyleInColumn(this ISheet sheet, ICellStyle style, bool createCells, int x)
//        {
//            sheet._SetStyleInColumnRange(style, createCells, x, x);
//        }

//        static public void _SetStyleInColumnRange(this ISheet sheet, ICellStyle style, bool createCells, int x1, int? x2 = null)
//        {
//            sheet.NewRange(1, x1, null, x2).SetStyle(style, createCells);
//        }

//        static public void _ReplaceStyleInColumnRange(this ISheet sheet, ICellStyle style1, ICellStyle style2, int x1, int? x2 = null)
//        {
//            sheet.NewRange(1, x1, null, x2).ReplaceStyle(style1, style2);
//        }

//        static public void _ClearStyleInColumnRange(this ISheet sheet, ICellStyle style, int x1, int? x2 = null)
//        {
//            sheet._ReplaceStyleInColumnRange(style, null, x1, x2);
//        }

//        /// <summary>
//        /// (!)Very slow on large data.
//        /// </summary>
//        /// <param name="x1"></param>
//        /// <param name="x2"></param>
//        /// <param name="padding">a character width</param>
//        static public void _AutosizeColumnsInRange(this ISheet sheet, int x1 = 1, int? x2 = null, float padding = 0)
//        {
//            if (x2 == null)
//                x2 = sheet._GetLastColumn();
//            for (int x = x1; x <= x2; x++)
//                sheet._AutosizeColumn(x, padding);
//        }

//        /// <summary>
//        /// (!)Very slow on large data.
//        /// </summary>
//        /// <param name="columnIs"></param>
//        /// <param name="padding">a character width</param>
//        static public void _AutosizeColumns(this ISheet sheet, IEnumerable<int> Xs, float padding = 0)
//        {
//            foreach (int y in Xs)
//                sheet._AutosizeColumn(y, padding);
//        }

//        /// <summary>
//        /// (!)Very slow on large data.
//        /// </summary>
//        /// <param name="x"></param>
//        /// <param name="padding">a character width</param>
//        static public void _AutosizeColumn(this ISheet sheet, int x, float padding = 0)
//        {
//            sheet._GetColumn(x).Autosize(padding);
//        }

//        static public IEnumerable<ICell> _GetCellsInColumn(this ISheet sheet, int x)
//        {
//            return sheet._GetColumn(x).GetCells();
//        }

//        /// <summary>
//        /// Safe against the API's one
//        /// </summary>
//        /// <param name="x"></param>
//        /// <param name="width">units of 1/256th of a character width</param>
//        static public void _SetColumnWidth(this ISheet sheet, int x, int width)
//        {
//            sheet._GetColumn(x).SetWidth(width);
//        }

//        /// <summary>
//        /// Safe against the API's one
//        /// </summary>
//        /// <param name="x"></param>
//        /// <param name="width">a character width</param>
//        static public void _SetColumnWidth(this ISheet sheet, int x, float width)
//        {
//            sheet._GetColumn(x).SetWidth(width);
//        }
//    }
//}