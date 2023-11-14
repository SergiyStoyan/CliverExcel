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
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public partial class Table
        {
            public Table(Excel excel)
            {
                Excel = excel;
                Sheet = Excel.Sheet;
                IRow headersRow = Sheet._GetRow(1, true);
                //empty-header columns are passed through
                List<Column> columns = headersRow._GetCells(true)
                    .Select(a => (header: a._GetValueAsString(), x: a._X()))
                    .Where(a => !string.IsNullOrEmpty(a.header))
                    .Select(a =>
                    {
                        Column c = new Column(a.header);
                        c.X = a.x;
                        c.Table = this;
                        return c;
                    })
                    .ToList();
                Columns = new ReadOnlyCollection<Column>(columns);
            }

            //public Table(Excel excel, SetColumnMode setColumnMode, params string[] headers) : this(excel, setColumnMode, (IEnumerable<string>)headers) { }

            //public Table(Excel excel, SetColumnMode setColumnMode, IEnumerable<string> headers) : this(excel)
            //{
            //    SetColumns(setColumnMode, headers);
            //}

            public Table(Excel excel, SetColumnMode setColumnMode, params Column[] columns) : this(excel, setColumnMode, (IEnumerable<Column>)columns) { }

            public Table(Excel excel, SetColumnMode setColumnMode, IEnumerable<Column> columns) : this(excel)
            {
                SetColumns(setColumnMode, columns);
            }

            readonly public ISheet Sheet;
            readonly public Excel Excel;

            /// <summary>
            /// Looks among the passed rows.
            /// (!)Rows and keys must belong to the same table.
            /// </summary>
            /// <param name="rows"></param>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            static public IEnumerable<IRow> FindRows(IEnumerable<IRow> rows, params Key[] keys)
            {
                return FindRows(rows, (IEnumerable<Key>)keys);
            }

            /// <summary>
            /// Key for seeking cell matches in a column.
            /// </summary>
            public class Key
            {
                public Column Column { get; internal set; }
                public int X { get { return Column.X; } }
                public Func<ICell, bool> IsValueMatch = null;

                /// <summary>
                /// Key that matches by isValueMatch().
                /// </summary>
                /// <param name="column"></param>
                /// <param name="isValueMatch"></param>
                /// <exception cref="Exception"></exception>
                public Key(Column column, Func<ICell, bool> isValueMatch)
                {
                    if (column.Table == null)
                        throw new Exception("Column is not initialized: Table is not set.");
                    Column = column;
                    IsValueMatch = isValueMatch;
                }

                /// <summary>
                /// Key that is equal to value.
                /// (!)Column must be listed in Table.Columns
                /// </summary>
                /// <param name="column"></param>
                /// <param name="value"></param>
                public Key(Column column, object value) : this(column, getIsValueMatch(value)) { }
                static Func<ICell, bool> getIsValueMatch(object value)
                {
                    string v = value.ToString();
                    return (c) => { return c?._GetValueAsString() == v; };
                }

                /// <summary>
                /// Key that matches by valueMatchRegex.
                /// (!)Column must be listed in Table.Columns
                /// </summary>
                /// <param name="column"></param>
                /// <param name="valueMatchRegex"></param>
                public Key(Column column, Regex valueMatchRegex) : this(column, getIsValueMatch(valueMatchRegex)) { }
                static Func<ICell, bool> getIsValueMatch(Regex valueMatchRegex)
                {
                    return (c) => { return valueMatchRegex.IsMatch(c._GetValueAsString()); };
                }

                /// <summary>
                /// Key that is equal to cell.Value.
                /// </summary>
                /// <param name="cell"></param>
                public Key(Cell cell) : this(cell.Column, cell.Value) { }

                /// <summary>
                /// Key that matches by isValueMatch().
                /// </summary>
                /// <param name="cell"></param>
                /// <param name="isValueMatch"></param>
                public Key(Cell cell, Func<ICell, bool> isValueMatch) : this(cell.Column, isValueMatch) { }

                /// <summary>
                /// Key that matches by valueMatchRegex().
                /// </summary>
                /// <param name="cell"></param>
                /// <param name="valueMatchRegex"></param>
                public Key(Cell cell, Regex valueMatchRegex) : this(cell.Column, getIsValueMatch(valueMatchRegex)) { }
            }

            public class Cell
            {
                public Column Column { get; internal set; }
                public object Value { get; internal set; }
                public int X { get { return Column.X; } }

                /// <summary>
                /// (!)Column must be listed in Table.Columns
                /// </summary>
                /// <param name="column"></param>
                /// <param name="value"></param>
                /// <exception cref="Exception"></exception>
                public Cell(Column column, object value)
                {
                    if (column.Table == null)
                        throw new Exception("Column is not initialized: Table is not set.");
                    Column = column;
                    Value = value;
                }

                //public NamedValue(Table table, string header, object value, Func<ICell, bool> isValueMatch = null)
                //{
                //    Column = table.GetColumn(header);
                //    Value = value;
                //    IsValueMatch = isValueMatch;
                //}

                //public NamedValue(Table table, Func<string, bool> isHeaderMatch, object value, Func<ICell, bool> isValueMatch = null)
                //{
                //    Column = table.GetColumn(isHeaderMatch);
                //    Value = value;
                //    IsValueMatch = isValueMatch;
                //}

                //public NamedValue(Table table, Regex headerRegex, object value, Func<ICell, bool> isValueMatch = null)
                //{
                //    Column = table.GetColumn(headerRegex);
                //    Value = value;
                //    IsValueMatch = isValueMatch;
                //}
            }

            /// <summary>
            /// Looks among the passed rows.
            /// (!)Rows and keys must belong to the same table.
            /// </summary>
            /// <param name="rows"></param>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            static public IEnumerable<IRow> FindRows(IEnumerable<IRow> rows, IEnumerable<Key> keys)
            {
                return rows.Where(a =>
                {
                    if (a == null)
                        return false;
                    foreach (var k in keys)
                    {
                        //if (a.Sheet != k.Column.Table.Sheet)
                        //    throw new Exception("Row[x=" + (a.RowNum + 1) + "] and key[X='" + k.X + "] belong to different sheets.");
                        if (!k.IsValueMatch(a.GetCell(k.X - 1)))
                            return false;
                    }
                    return true;
                });
            }

            /// <summary>
            /// (!)Re-reads the sheet on every call which can be slow.
            /// </summary>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            public IEnumerable<IRow> FindDataRows(params Key[] keys)
            {
                return FindRows(GetDataRows(RowScope.WithCells), keys);
            }

            public IEnumerable<IRow> GetDataRows(RowScope rowScope)
            {
                return Sheet._GetRows(rowScope).Skip(1);
            }

            //public IRow AppendRow<T>(IEnumerable<T> values)
            //{
            //    IRow r = Sheet._AppendRow(values);
            //    setColumnStyles(r);

            //    //if (cachedDataRows != null)
            //    //    cachedDataRows.Add(r);

            //    return r;
            //}

            void setColumnStyles(IRow row)
            {
                foreach (Column c in Columns)
                    row._GetCell(c.X, true).CellStyle = c.DataStyle;
            }

            void setColumnTypes(IRow row)
            {
                foreach (Column c in Columns)
                    if (c.DataType != null)
                        row._GetCell(c.X, true).SetCellType(c.DataType.Value);
            }

            //public IRow AppendRow(params string[] values)
            //{
            //    return AppendRow((IEnumerable<string>)values);
            //}

            public IRow AppendRow(IEnumerable<Cell> cells)
            {
                IRow r = WriteRow(Sheet._GetLastRow(LastRowCondition.HasCells, false) + 1, cells);
                return r;
            }

            public IRow AppendRow(params Cell[] cells)
            {
                return AppendRow((IEnumerable<Cell>)cells);
            }

            //public IRow InsertRow<T>(int y, IEnumerable<T> values = null)
            //{
            //    IRow r = Sheet._InsertRow(y, values);
            //    setColumnStyles(r);

            //    //if (cachedDataRows != null)
            //    //    cachedDataRows.Insert(r.RowNum, r);

            //    return r;
            //}

            //public IRow InsertRow(int y, params string[] values)
            //{
            //    return InsertRow(y, (IEnumerable<string>)values);
            //}

            public IRow InsertRow(int y, params Cell[] cells)
            {
                return InsertRow(y, (IEnumerable<Cell>)cells);
            }

            public IRow InsertRow(int y, IEnumerable<Cell> cells)
            {
                int lastRowY = Sheet._GetLastRow(LastRowCondition.HasCells, false);
                if (y <= lastRowY)
                    Sheet.ShiftRows(y - 1, lastRowY - 1, 1);
                IRow r = WriteRow(y, cells);
                return r;
            }

            //public IRow WriteRow<T>(int y, IEnumerable<T> values = null)
            //{
            //    IRow r = Sheet._WriteRow(y, values);
            //    return r;
            //}

            //public IRow WriteRow(int y, params string[] values)
            //{
            //    return WriteRow(y, (IEnumerable<string>)values);
            //}

            public IRow WriteRow(int y, IEnumerable<Cell> cells)
            {
                IRow r = Sheet.GetRow(y - 1);
                if (r == null)
                {
                    r = Sheet.CreateRow(y - 1);
                    setColumnStyles(r);
                    setColumnTypes(r);
                }
                foreach (var cell in cells)
                {
                    //if (r.Sheet != cell.Column.Table.Sheet)
                    //    throw new Exception("Row[x=" + (r.RowNum + 1) + "] and cell[X='" + cell.X + "] belong to different sheets.");
                    var c = r._GetCell(cell.X, true);
                    c._SetValue(cell.Value);
                }
                return r;
            }

            public IRow WriteRow(int y, params Cell[] cells)
            {
                return WriteRow(y, (IEnumerable<Cell>)cells);
            }

            public IRow RemoveRow(int y, bool shiftRowsBelow)
            {
                return Sheet._RemoveRow(y, shiftRowsBelow);
            }

            public void MoveRow(int y1, int y2)
            {
                Sheet._MoveRow(y1, y2);
            }

            ///// <summary>
            ///// (!)Seeks the column each call.
            ///// </summary>
            ///// <param name="row"></param>
            ///// <param name="header"></param>
            ///// <param name="create"></param>
            ///// <returns></returns>
            //public ICell GetCell(IRow row, string header, bool create)
            //{
            //    return row._GetCell(GetColumn(header).X, create);
            //}

            public ICell GetCell(IRow row, Column column, bool create)
            {
                return row._GetCell(column.X, create);
            }

            public ICell GetCell(int y, Column column, bool create)
            {
                return Sheet._GetCell(y, column.X, create);
            }

            public void SetStyles(IRow row, params ICellStyle[] styles)
            {
                row._SetStyles(1, styles);
            }

            public void SetStyles(IRow row, IEnumerable<ICellStyle> styles)
            {
                SetStyles(row, styles.ToArray());
            }

            public void Save(string file = null)
            {
                Excel.Save(file);
            }
        }
    }
}