//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace Cliver
{
    public partial class Excel
    {
        public void ShiftCellsRight(int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftCellsRight(x1, y1, y2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsLeft(int x1, int y1, int y2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftCellsLeft(x1, y1, y2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsDown(int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftCellsDown(y1, x1, x2, shift, onFormulaCellMoved);
        }

        public void ShiftCellsUp(int y1, int x1, int x2, int shift, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.ShiftCellsUp(y1, x1, x2, shift, onFormulaCellMoved);
        }

        public void CopyCell(int fromCellY, int fromCellX, int toCellY, int toCellX)
        {
            Sheet.CopyCell(fromCellY, fromCellX, toCellY, toCellX);
        }

        public void MoveCell(int fromCellY, int fromCellX, int toCellY, int toCellX, Action<ICell> onFormulaCellMoved = null)
        {
            Sheet.MoveCell(fromCellY, fromCellX, toCellY, toCellX, onFormulaCellMoved);
        }

        public ICell GetCell(int y, int x, bool createCell)
        {
            return Sheet.GetCell(y, x, createCell);
        }

        public ICell GetCell(string address, bool createCell)
        {
            return Sheet.GetCell(address, createCell);
        }

        //public void HighlightCell(int y, int x, ICellStyle style, Color color)
        //{
        //    GetCell(y, x, true).Highlight(style, color);
        //}

        //public void Highlight(ICell cell, ICellStyle style, Color color)
        //{
        //    cell.Highlight(style, color);
        //}

        public void SetLink(int y, int x, Uri uri)
        {
            Sheet.SetLink(y, x, uri);
        }

        public Uri GetLink(int y, int x)
        {
            return Sheet.GetLink(y, x);
        }

        public string GetValueAsString(int y, int x, bool allowNull = false)
        {
            return Sheet.GetValueAsString(y, x, allowNull);
        }

        public object GetValue(int y, int x)
        {
            return Sheet.GetValue(y, x);
        }

        public void SetValue(int y, int x, object value)
        {
            Sheet.SetValue(y, x, value);
        }

        /// <summary>
        /// (!)Never returns NULL.
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        public string this[int y, int x]
        {
            get
            {
                return Sheet.GetValueAsString(y, x, false);
            }
            set
            {
                Sheet.SetValue(y, x, value);
            }
        }

        public IEnumerable<Image> GetImages(int y, int x)
        {
            return Sheet.GetImages(y, x);
        }

        public void CreateDropdown<T>(int y, int x, IEnumerable<T> values, T value, bool allowBlank = true)
        {
            Sheet.CreateDropdown(y, x, values, value, allowBlank);
        }
    }
}