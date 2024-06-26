//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using static Cliver.Excel;

namespace Cliver
{
    static public partial class ExcelTableExtensions
    {
        static public void _ShiftCellsRight(this IRow row, Excel.Table.Column c1, int shift, CopyCellMode copyCellMode = null)
        {
            row._ShiftCellsRight(c1.X, shift, copyCellMode);
        }

        static public void _ShiftCellsLeft(this IRow row, Excel.Table.Column c1, int shift, CopyCellMode copyCellMode = null)
        {
            row._ShiftCellsLeft(c1.X, shift, copyCellMode);
        }

        static public ICell _GetCell(this IRow row, Excel.Table.Column c, bool createCell)
        {
            return row._GetCell(c.X, createCell);
        }

        static public IEnumerable<ICell> _GetCellsInRange(this IRow row, CellScope cellScope, Excel.Table.Column c1, Excel.Table.Column c2)
        {
            if (c1 == null)
                return c2 == null ? row._GetCellsInRange(cellScope) : row._GetCellsInRange(cellScope, 1, c2.X);
            return c2 == null ? row._GetCellsInRange(cellScope, c1.X) : row._GetCellsInRange(cellScope, c1.X, c2.X);
        }

        /// <summary> 
        /// Value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        static public object _GetValue(this IRow row, Excel.Table.Column c)
        {
            return row._GetValue(c.X);
        }

        /// <summary> 
        /// Set value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        static public void _SetValue(this IRow row, Excel.Table.Column c, object value)
        {
            row._SetValue(c.X, value);
        }

        /// <summary>
        /// Value of the specified cell.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="c"></param>
        /// <param name="stringMode"></param>
        /// <returns></returns>
        static public string _GetValueAsString(this IRow row, Excel.Table.Column c, StringMode stringMode = DefaultStringMode)
        {
            return row._GetValueAsString(c.X, stringMode);
        }

        /// <summary>
        /// Images anchored in the specified cell coordinates. The cell may not exist.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        static public IEnumerable<Excel.Image> _GetImages(this IRow row, Excel.Table.Column c)
        {
            return row._GetImages(c.X);
        }

        static public string _GetLink(this IRow row, Excel.Table.Column c)
        {
            return row?._GetLink(c.X);
        }

        static public void _SetLink(this IRow row, Excel.Table.Column c, string link)
        {
            row._SetLink(c.X, link);
        }

        static public void _SetStyles(this IRow row, IEnumerable<Excel.Table.Style> styles)
        {
            foreach (Excel.Table.Style s in styles.Where(a => a.Value != null))
                row._GetCell(s.Column.X, true).CellStyle = s.Value;
        }

        static public void _SetStyles(this IRow row, params Excel.Table.Style[] styles)
        {
            _SetStyles(row, (IEnumerable<Excel.Table.Style>)styles);
        }
    }
}
