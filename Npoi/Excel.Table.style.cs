//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public partial class Table
        {
            public class Style
            {
                public Column Column { get; internal set; }
                public ICellStyle Value { get; internal set; }

                /// <summary>
                /// (!)Unregistered style will be registered.
                /// (!)Column must be listed in Table.Columns.
                /// </summary>
                /// <param name="column"></param>
                /// <param name="style"></param>
                /// <exception cref="Exception"></exception>
                public Style(Column column, ICellStyle style = null)
                {
                    if (column.Table == null)
                        throw new Exception("Column is not initialized: Table is not set.");
                    Column = column;
                    if (style == null)
                        style = column.DataStyle;
                    if (style.Index < 0)
                        style = column.Table.Excel.Workbook._GetRegisteredStyle(style);
                    Value = style;
                }

                public Cell NewCell(object value, CellType? type = null)
                {
                    return new Cell(this, value, type);
                }
            }

            static public void SetStyles(IRow row, IEnumerable<Style> styles)
            {
                row._SetStyles(styles);
            }

            public void SetStyles(int y, IEnumerable<Style> styles)
            {
                Sheet._GetRow(y, true)._SetStyles(styles);
            }

            static public void SetStyles(IRow row, params Style[] styles)
            {
                row._SetStyles((IEnumerable<Style>)styles);
            }

            public void SetStyles(int y, params Style[] styles)
            {
                Sheet._GetRow(y, true)._SetStyles(styles);
            }

            /// <summary>
            /// Sets the row with styles that are obtained by altering the row's existing styles.
            /// This function takes care about caching and registering all the styles needed.
            /// Based on Excel.StyleCache.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="alterationKey">a key for the given style alteration, e.g. changing to a font. (!)It must be unique for all the planned alterations.</param>
            /// <param name="updateStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            public void SetAlteredStyles<T>(IRow row, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                foreach (Column column in Columns)
                    row._GetCell(column, true)._SetAlteredStyle(alterationKey, alterStyle, reuseUnusedStyle);
            }

            public void SetAlteredStyles<T>(int y, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                SetAlteredStyles(Sheet._GetRow(y, true), alterationKey, alterStyle, reuseUnusedStyle);
            }

            public void SetAlteredStyle<T>(IRow row, Column column, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, bool reuseUnusedStyle = false) where T : Excel.StyleCache.IKey
            {
                row._GetCell(column, true)._SetAlteredStyle(alterationKey, alterStyle, reuseUnusedStyle);
            }

            public void SetAlteredStyle<T>(int y, Column column, T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle) where T : Excel.StyleCache.IKey
            {
                SetAlteredStyle(Sheet._GetRow(y, true), column, alterationKey, alterStyle);
            }

            public void SetStyle(IRow row, ICellStyle style, RowStyleMode rowStyleMode)
            {
                if (rowStyleMode.HasFlag(RowStyleMode.Row))
                    row.RowStyle = style;
                if (rowStyleMode.HasFlag(RowStyleMode.ExistingCells))
                    foreach (Column column in Columns)
                        row._GetCell(column, true).CellStyle = style;
                else if (rowStyleMode.HasFlag(RowStyleMode.NoGapCells))
                    for (int? x = Columns.LastOrDefault()?.X; x > 0; x--)
                        row._GetCell(x.Value, true).CellStyle = style;
            }
        }
    }
}