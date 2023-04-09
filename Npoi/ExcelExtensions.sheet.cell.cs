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
//        static public void SetLink(this ISheet sheet, int y, int x, Uri uri)
//        {
//            sheet.GetCell(y, x, true).SetLink(uri);
//        }

//        static public Uri GetLink(this ISheet sheet, int y, int x)
//        {
//            return sheet.GetCell(y, x, false)?.GetLink();
//        }

//        static public void ShiftCellsRight(this ISheet sheet, int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            for (int y = y1; y <= y2; y++)
//            {
//                for (int x = sheet.GetLastNotEmptyColumnInRow(y); x >= x1; x--)
//                    sheet.MoveCell(y, x, y, x + shift, onFormulaCellMoved);
//                sheet.GetCell(y, x1, false)?.SetBlank();
//            }
//        }

//        static public void ShiftCellsLeft(this ISheet sheet, int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            for (int y = y1; y <= y2; y++)
//            {
//                for (int x = 1; x <= x1; x++)
//                    sheet.MoveCell(y, x, y, x - shift, onFormulaCellMoved);
//                sheet.GetCell(y, x1, false)?.SetBlank();
//            }
//        }

//        static public void ShiftCellsDown(this ISheet sheet, int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            for (int x = x1; x <= x2; x++)
//            {
//                for (int y = sheet.GetLastNotEmptyRowInColumn(x); y >= y1; y--)
//                    sheet.MoveCell(y, x, y + shift, x, onFormulaCellMoved);
//                sheet.GetCell(y1, x, false)?.SetBlank();
//            }
//        }

//        static public void ShiftCellsUp(this ISheet sheet, int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
//        {
//            for (int x = x1; x <= x2; x++)
//            {
//                for (int y = 1; y <= y1; y++)
//                    sheet.MoveCell(y, x, y - shift, x, onFormulaCellMoved);
//                sheet.GetCell(y1, x, false)?.SetBlank();
//            }
//        }

//        static public void CopyCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX)
//        {
//            ICell sourceCell = sheet.GetCell(fromCellY, fromCellX, false);
//            sourceCell.Copy(toCellY, toCellX);
//        }

//        static public string GetValueAsString(this ISheet sheet, int y, int x, bool allowNull = false)
//        {
//            ICell c = sheet.GetCell(y, x, false);
//            return c?.GetValueAsString(allowNull);
//        }

//        static public object GetValue(this ISheet sheet, int y, int x)
//        {
//            ICell c = sheet.GetCell(y, x, false);
//            return c?.GetValue();
//        }

//        static public void SetValue(this ISheet sheet, int y, int x, object value)
//        {
//            ICell c = sheet.GetCell(y, x, true);
//            c.SetValue(value);
//        }

//        static public void MoveCell(this ISheet sheet, int fromCellY, int fromCellX, int toCellY, int toCellX, Action<ICell> onFormulaCellMoved = null)
//        {
//            ICell fromCell = sheet.GetCell(fromCellY, fromCellX, false);
//            fromCell.Move(toCellY, toCellX, onFormulaCellMoved);
//        }

//        static public ICell GetCell(this ISheet sheet, int y, int x, bool createCell)
//        {
//            IRow r = sheet.GetRow(y, createCell);
//            if (r == null)
//                return null;
//            return r.GetCell(x, createCell);
//        }

//        static public ICell GetCell(this ISheet sheet, string address, bool createCell)
//        {
//            var cs = GetCoordinates(address);
//            IRow r = sheet.GetRow(cs.Y, createCell);
//            if (r == null)
//                return null;
//            return r.GetCell(cs.X, createCell);
//        }

//        static public void RemoveCell(this ISheet sheet, int y, int x)
//        {
//            IRow r = sheet.GetRow(y);
//            if (r == null)
//                return;
//            ICell c = r.GetCell(x);
//            if (c == null)
//                return;
//            r.RemoveCell(c);
//        }

//        static internal Range getMergedRange(this ISheet sheet, int y, int x)
//        {
//            foreach (var mr in sheet.MergedRegions)
//                if (mr.IsInRange(y - 1, x - 1))
//                    return new Range(sheet, mr.FirstRow + 1, mr.FirstColumn + 1, mr.LastRow + 1, mr.LastColumn + 1);
//            return null;
//        }

//        static public void CreateDropdown<T>(this ISheet sheet, int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
//        {
//            sheet.CreateDropdown(y, x, values, value, allowBlank);
//        }

//    }
//}