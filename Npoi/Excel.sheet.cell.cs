//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace Cliver
{
    public partial class Excel
    {
        public void SetComment(int y, int x, string comment, string author = null, IClientAnchor anchor = null)
        {
            Sheet._SetComment(y, x, comment, author, anchor);
        }

        public void _AppendOrSetComment(int y, int x, string comment, string author = null, string separator = "\r\n\r\n", IClientAnchor anchor = null)
        {
            Sheet._AppendOrSetComment(y, x, comment, author, separator, anchor);
        }

        public static string LinkEmptyValueFiller = "           ";
        public void SetLink(int y, int x, string link)
        {
            Sheet._SetLink(y, x, link);
        }

        public string GetLink(int y, int x)
        {
            return Sheet._GetLink(y, x);
        }

        public delegate void OnFormulaCellMoved(ICell fromCell, ICell toCell);
        public void ShiftCellsRight(int x1, int y1, int y2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftCellsRight(x1, y1, y2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsLeft(int x1, int y1, int y2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftCellsLeft(x1, y1, y2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsDown(int y1, int x1, int x2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftCellsDown(y1, x1, x2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsUp(int y1, int x1, int x2, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftCellsUp(y1, x1, x2, shift, onFormulaCellMoved);
        }

        public ICell CopyCell(int fromCellY, int fromCellX, int toCellY, int toCellX)
        {
            return Sheet._CopyCell(fromCellY, fromCellX, toCellY, toCellX);
        }

        public string GetValueAsString(int y, int x, bool allowNull = false)
        {
            return Sheet._GetValueAsString(y, x, allowNull);
        }

        public object GetValue(int y, int x)
        {
            return Sheet._GetValue(y, x);
        }

        public void SetValue(int y, int x, object value)
        {
            Sheet._SetValue(y, x, value);
        }

        public ICell MoveCell(int fromCellY, int fromCellX, int toCellY, int toCellX, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
        {
            return Sheet._MoveCell(fromCellY, fromCellX, toCellY, toCellX, onFormulaCellMoved, toSheet);
        }

        public ICell GetCell(int y, int x, bool createCell)
        {
            return Sheet._GetCell(y, x, createCell);
        }

        public ICell GetCell(string address, bool createCell)
        {
            return Sheet._GetCell(address, createCell);
        }

        public void RemoveCell(int y, int x)
        {
            Sheet._RemoveCell(y, x);
        }

        public void UpdateFormulaRange(int y, int x, int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
        {
            Sheet._UpdateFormulaRange(y, x, rangeY1Shift, rangeX1Shift, rangeY2Shift, rangeX2Shift);
        }

        public void ClearMerging(int y, int x)
        {
            Sheet._ClearMerging(y, x);
        }

        public void CreateDropdown<T>(int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            Sheet._CreateDropdown(y, x, values, value, allowBlank);
        }

        public void AddImage(Image image)
        {
            Sheet._AddImage(image);
        }

        public Range GetMergedRange(int y, int x)
        {
            return Sheet._GetMergedRange(y, x);
        }

        public IEnumerable<Image> GetImages(int y, int x)
        {
            return Sheet._GetImages(y, x);
        }
    }
}