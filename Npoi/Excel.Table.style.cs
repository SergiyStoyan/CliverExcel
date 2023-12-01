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
                        style = column.Table.Excel.GetRegisteredStyle(style);
                    Value = style;
                }

                public Cell NewCell(object value, CellType? type = null)
                {
                    return new Cell(this, value, type);
                }
            }

            public void SetStyles(IRow row, IEnumerable<Style> styles)
            {
                row._SetStyles(styles);
            }

            public void SetStyles(IRow row, params Style[] styles)
            {
                row._SetStyles((IEnumerable<Style>)styles);
            }

            /// <summary>
            /// Sets the row with styles that are obtained by altering the row's existing styles.
            /// This function takes care about caching and registering all the styles needed.
            /// Based on Excel.StyleCache.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="alterationKey">a key for the given style alteration, e.g. changing to a font. (!)It must be unique for all the planned alterations.</param>
            /// <param name="updateStyle">performs style alteration. (!)The passed in style is unregistered and must remain so.</param>
            public void SetStyles(IRow row, Excel.StyleCache.IKey alterationKey, Action<ICellStyle> alterStyle)
            {
                foreach (Column column in Columns)
                    Excel.SetStyle(row._GetCell(column, true), alterationKey, alterStyle);
            }
        }
    }
}