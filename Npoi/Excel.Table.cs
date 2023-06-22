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

            public class Key
            {
                public Column Column { get; internal set; }
                public int X { get { return Column.X; } }
                public Func<ICell, bool> IsValueMatch = null;

                public Key(Column column, Func<ICell, bool> isValueMatch)
                {
                    if (column.Table == null)
                        throw new Exception("Column is not initialized: no Table set.");
                    Column = column;
                    IsValueMatch = isValueMatch;
                }

                public Key(Column column, object value) : this(column, getIsValueMatch(value)) { }
                static Func<ICell, bool> getIsValueMatch(object value)
                {
                    string v = value.ToString();
                    return (c) => { return c?._GetValueAsString() == v; };
                }

                public Key(Cell cell) : this(cell.Column, cell.Value) { }

                public Key(Cell cell, Func<ICell, bool> isValueMatch) : this(cell.Column, isValueMatch) { }

                public Key(Cell cell, Regex isValueMatchRegex) : this(cell.Column, getIsValueMatch(isValueMatchRegex)) { }
                static Func<ICell, bool> getIsValueMatch(Regex isValueMatchRegex)
                {
                    return (c) => { return isValueMatchRegex.IsMatch(c._GetValueAsString()); };
                }
            }

            public class Cell
            {
                public Column Column { get; internal set; }
                public object Value { get; internal set; }
                public int X { get { return Column.X; } }

                public Cell(Column column, object value)
                {
                    if (column.Table == null)
                        throw new Exception("Column is not initialized: no Table set.");
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
                        if (!k.IsValueMatch(a.GetCell(k.X - 1)))
                            return false;
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
                return FindRows(Sheet._GetRows(RowScope.WithCells).Skip(1), keys);
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

            //public IRow InsertFullRow(int y, params NamedValue[] values)
            //{
            //    return InsertFullRow(y, (IEnumerable<NamedValue>)values);
            //}

            //public IRow InsertFullRow(int y, IEnumerable<NamedValue> namedValues)
            //{
            //    return InsertRow(y, (IEnumerable<NamedValue>)values);
            //}

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

                    //if (cachedDataRows != null)
                    //    cachedDataRows.Insert(r.RowNum, r);
                }
                foreach (var cell in cells)
                {
                    var c = r._GetCell(cell.X, true);
                    c._SetValue(cell.Value);
                }
                return r;
            }

            public IRow WriteRow(int y, params Cell[] cells)
            {
                return WriteRow(y, (IEnumerable<Cell>)cells);
            }

            public IRow RemoveRow(int y, bool shiftRemainingRows)
            {
                return Sheet._RemoveRow(y, shiftRemainingRows);

                //if (cachedDataRows != null && r != null)
                //    cachedDataRows.Remove(r);
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