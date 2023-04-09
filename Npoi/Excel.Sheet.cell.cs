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
    public partial class Sheet
    {
         public void SetLink( int y, int x, Uri uri)
        {
            GetCell(y, x, true).SetLink(uri);
        }

         public Uri GetLink( int y, int x)
        {
            return GetCell(y, x, false)?.GetLink();
        }

         public void ShiftCellsRight( int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
            {
                for (int x = GetLastNotEmptyColumnInRow(y); x >= x1; x--)
                    MoveCell(y, x, y, x + shift, onFormulaCellMoved);
                GetCell(y, x1, false)?._.SetBlank();
            }
        }

         public void ShiftCellsLeft( int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int y = y1; y <= y2; y++)
            {
                for (int x = 1; x <= x1; x++)
                    MoveCell(y, x, y, x - shift, onFormulaCellMoved);
                GetCell(y, x1, false)?._.SetBlank();
            }
        }

         public void ShiftCellsDown( int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
            {
                for (int y = GetLastNotEmptyRowInColumn(x); y >= y1; y--)
                    MoveCell(y, x, y + shift, x, onFormulaCellMoved);
                GetCell(y1, x, false)?._.SetBlank();
            }
        }

         public void ShiftCellsUp( int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            for (int x = x1; x <= x2; x++)
            {
                for (int y = 1; y <= y1; y++)
                    MoveCell(y, x, y - shift, x, onFormulaCellMoved);
                GetCell(y1, x, false)?._.SetBlank();
            }
        }

         public void CopyCell( int fromCellY, int fromCellX, int toCellY, int toCellX)
        {
            Cell sourceCell = GetCell(fromCellY, fromCellX, false);
            sourceCell.Copy(toCellY, toCellX);
        }

         public string GetValueAsString( int y, int x, bool allowNull = false)
        {
            Cell c = GetCell(y, x, false);
            return c?.GetValueAsString(allowNull);
        }

         public object GetValue( int y, int x)
        {
            Cell c = GetCell(y, x, false);
            return c?.GetValue();
        }

         public void SetValue( int y, int x, object value)
        {
            Cell c = GetCell(y, x, true);
            c.SetValue(value);
        }

         public void MoveCell( int fromCellY, int fromCellX, int toCellY, int toCellX, Action<Cell> onFormulaCellMoved = null)
        {
            Cell fromCell = GetCell(fromCellY, fromCellX, false);
            fromCell.Move(toCellY, toCellX, onFormulaCellMoved);
        }

         public Cell GetCell( int y, int x, bool createCell)
        {
            Row r = GetRow(y, createCell);
            if (r == null)
                return null;
            return r.GetCell(x, createCell);
        }

         public ICell GetCell( string address, bool createCell)
        {
            var cs = GetCoordinates(address);
            IRow r = GetRow(cs.Y, createCell);
            if (r == null)
                return null;
            return r.GetCell(cs.X, createCell);
        }

         public void RemoveCell( int y, int x)
        {
            IRow r = GetRow(y);
            if (r == null)
                return;
            ICell c = r.GetCell(x);
            if (c == null)
                return;
            r.RemoveCell(c);
        }

         internal Range getMergedRange( int y, int x)
        {
            foreach (var mr in MergedRegions)
                if (mr.IsInRange(y - 1, x - 1))
                    return new Range(sheet, mr.FirstRow + 1, mr.FirstColumn + 1, mr.LastRow + 1, mr.LastColumn + 1);
            return null;
        }

         public void CreateDropdown<T>( int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            CreateDropdown(y, x, values, value, allowBlank);
        }

    }
}