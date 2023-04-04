//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;

namespace Cliver
{
    public partial class Excel
    {
        public void ShiftCellsDown(int cellsY, int firstCellX, int lastCellX, int rowCount, Action<ICell> updateFormula = null)
        {
            for (int x = firstCellX; x <= lastCellX; x++)
            {
                for (int y = GetLastNotEmptyRowInColumn(x); y >= cellsY; y--)
                {
                    CopyCell(y, x, y + rowCount, x);
                    if (updateFormula == null)
                        continue;
                    ICell formulaCell = GetCell(y + rowCount, x, false);
                    if (formulaCell?.CellType != CellType.Formula)
                        continue;
                    updateFormula(formulaCell);
                }
                GetCell(cellsY, x, false)?.SetBlank();
            }
        }

        public void CopyCell(ICell source, ICell destination)
        {
            destination.SetBlank();
            destination.SetCellType(source.CellType);
            destination.CellStyle = source.CellStyle;
            destination.CellComment = source.CellComment;
            destination.Hyperlink = source.Hyperlink;
            switch (source.CellType)
            {
                case CellType.Formula:
                    destination.CellFormula = source.CellFormula;
                    break;
                case CellType.Numeric:
                    destination.SetCellValue(source.NumericCellValue);
                    break;
                case CellType.String:
                    destination.SetCellValue(source.StringCellValue);
                    break;
                case CellType.Boolean:
                    destination.SetCellValue(source.BooleanCellValue);
                    break;
                case CellType.Error:
                    destination.SetCellErrorValue(source.ErrorCellValue);
                    break;
                case CellType.Blank:
                    destination.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown cell type: " + source.CellType);
            }
        }

        public ICell CopyCell(ICell sourceCell, int destinationY, int destinationX)
        {
            if (sourceCell == null)
            {
                IRow destinationRow = GetRow(destinationY, false);
                if (destinationRow == null)
                    return null;
                ICell destinationCell = destinationRow.GetCell(destinationX, false);
                if (destinationCell == null)
                    return destinationCell;
                destinationRow.RemoveCell(destinationCell);
                return destinationCell;
            }
            else
            {
                ICell destinationCell = GetCell(destinationY, destinationX, true);
                CopyCell(sourceCell, destinationCell);
                return destinationCell;
            }
        }

        public void MoveCell(ICell sourceCell, int destinationY, int destinationX, Action<ICell> onFormulaCellMoved = null)
        {
            ICell destinationCell = CopyCell(sourceCell, destinationY, destinationX);
            if (sourceCell != null)
                sourceCell.Row.RemoveCell(sourceCell);
            if (destinationCell?.CellType == CellType.Formula)
                onFormulaCellMoved?.Invoke(destinationCell);
        }

        public void CopyCell(int sourceY, int sourceX, int destinationY, int destinationX)
        {
            ICell sourceCell = GetCell(sourceY, sourceX, false);
            CopyCell(sourceCell, destinationY, destinationX);
        }

        public void MoveCell(int sourceY, int sourceX, int destinationY, int destinationX, Action<ICell> onFormulaCellMoved = null)
        {
            ICell sourceCell = GetCell(sourceY, sourceX, false);
            MoveCell(sourceCell, destinationY, destinationX, onFormulaCellMoved);
        }

        public ICell GetCell(int y, int x, bool create)
        {
            IRow r = GetRow(y, create);
            if (r == null)
                return null;
            return r.GetCell(x, create);
        }

        public ICell GetCell(string address, bool create)
        {
            var cs = GetCoordinates(address);
            IRow r = GetRow(cs.Y, create);
            if (r == null)
                return null;
            return r.GetCell(cs.X, create);
        }

        //public void HighlightCell(int y, int x, ICellStyle style, Color color)
        //{
        //    GetCell(y, x, true).Highlight(style, color);
        //}

        //public void Highlight(ICell cell, ICellStyle style, Color color)
        //{
        //    cell.Highlight(style, color);
        //}
    }
}