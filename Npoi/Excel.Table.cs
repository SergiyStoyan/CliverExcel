//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
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
        public class Table
        {
            //public Table(ISheet sheet)
            //{
            //    Sheet = sheet;
            //    IRow headersRow = Sheet._GetRow(1, true);
            //    headers = new ReadOnlyCollection<string>(headersRow._GetCells(true).Select(a => a._GetValueAsString()).ToList());
            //}

            //public Table(ISheet sheet, params string[] headers)
            //{
            //    Sheet = sheet;
            //    Headers = new ReadOnlyCollection<string>(headers.ToList());
            //}

            public Table(Excel excel)
            {
                Excel = excel;
                Sheet = Excel.Sheet;
                IRow headersRow = Sheet._GetRow(1, true);
                int emptyCount = 0;
                var hs = headersRow._GetCells(true).Select(a => a._GetValueAsString()).TakeWhile(a => !string.IsNullOrWhiteSpace(a) || ++emptyCount < 2).ToList();
                SetColumns(hs);
            }

            public Table(Excel excel, params string[] headers)
            {
                Excel = excel;
                Sheet = Excel.Sheet;
                SetColumns(headers);
            }

            public Table(Excel excel, params Column[] columns)
            {
                Excel = excel;
                Sheet = Excel.Sheet;
                SetColumns(columns);
            }

            readonly public ISheet Sheet;
            readonly public Excel Excel;

            public ReadOnlyCollection<Column> Columns { get; private set; }

            public void SetColumns(params string[] headers)
            {
                SetColumns((IEnumerable<string>)headers);
            }

            public void SetColumns(IEnumerable<string> headers)
            {
                SetColumns(headers.Select(a => new Column(a)));
            }

            public void SetColumns(params Column[] columns)
            {
                SetColumns((IEnumerable<Column>)columns);
            }

            public void SetColumns(IEnumerable<Column> columns)
            {
                var duplicates = columns.GroupBy(a => a.Header).Where(a => a.Count() > 1).Select(a => "'" + a.Key + "'").ToList();
                if (duplicates.Count > 0)
                    throw new Exception("Columns duplicated: " + string.Join(", ", duplicates));
                Columns = new ReadOnlyCollection<Column>(columns?.ToList());
                Columns?.Select((a, i) => (column: a, x: i + 1)).ForEach(a => { a.column.X = a.x; });
                Sheet._GetRow(1, true)._Write(Columns?.Select((a, i) => a.Header));
                //headers2Column.Clear();
                //Columns?.ForEach(a => headers2Column[a.Header] = a);
            }

            //Dictionary<string, Column> headers2Column = new Dictionary<string, Column>();

            public Column GetColumn(string header, bool exceptionIfNotFound = true)
            {
                //headers2Column.TryGetValue(header, out Column column);
                return GetColumn((v) => { return v == header; }, exceptionIfNotFound);
            }

            public Column GetColumn(Regex headerRegex, bool exceptionIfNotFound = true)
            {
                return GetColumn((v) => { return headerRegex.IsMatch(v); }, exceptionIfNotFound);
            }

            public Column GetColumn(Func<string, bool> IsHeaderMatch, bool exceptionIfNotFound = true)
            {
                var c = Columns.FirstOrDefault(a => IsHeaderMatch(a.Header));
                if (c == null && exceptionIfNotFound)
                    throw new Exception("Column was not found.");
                return c;
            }

            public class Column
            {
                //public class Columns : ReadOnlyCollection<Column>
                //{
                //    public Column this[string header]
                //    {
                //        get
                //        {
                //            return this.FirstOrDefault(a => a.Header == header);
                //        }
                //    }
                //}
                public readonly string Header;
                public int X { get; internal set; } = -1;
                public ICellStyle DataStyle;
                //public Func<string, bool> IsMatch = (v)=> { return v == Header; } ;

                internal Column(string header, int x, ICellStyle dataStyle)
                {
                    Header = header;
                    X = x;
                    DataStyle = dataStyle;
                }

                /// <summary>
                /// Used only to write new headers
                /// </summary>
                /// <param name="header"></param>
                /// <param name="style"></param>
                public Column(/*Table table, */string header, ICellStyle dataStyle = null/*, Func<ICell, bool> isMatch = null*/)
                {
                    Header = header;
                    DataStyle = dataStyle;
                }
            }

            ///// <summary>
            ///// Alias for UseCachedDataRows=true
            ///// </summary>
            //public void ReloadCachedDataRows()
            //{
            //    UseCachedDataRows = true;
            //}

            //List<IRow> getDataRows()
            //{
            //    return Sheet._GetRowsInRange(RowScope.IncludeNull, 2).ToList();
            //}

            ///// <summary>
            ///// Saves performance when FindRows() is called more than once. Otherwise, keep it switched off.
            ///// (!)Every time it is set, the cache is reloaded which should be done if the sheet was edited outside this class.
            ///// </summary>
            //public bool UseCachedDataRows
            //{
            //    set
            //    {
            //        cachedDataRows = value ? getDataRows() : null;
            //    }
            //    get
            //    {
            //        return cachedDataRows != null;
            //    }
            //} 
            //List<IRow> cachedDataRows = null;

            /// <summary>
            /// Looks among the passed rows.
            /// </summary>
            /// <param name="rows"></param>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            static public IEnumerable<IRow> FindRows(IEnumerable<IRow> rows, params NamedValue[] rowKeys)
            {
                return FindRows(rows, (IEnumerable<NamedValue>)rowKeys);
            }

            /// <summary>
            /// Looks among the passed rows.
            /// </summary>
            /// <param name="rows"></param>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            static public IEnumerable<IRow> FindRows(IEnumerable<IRow> rows, IEnumerable<NamedValue> rowKeys)
            {
                foreach (NamedValue rk in rowKeys)
                {
                    if (rk.IsValueMatch == null)
                    {
                        string valueAsString = rk.Value?.ToString();
                        rk.IsValueMatch = (ICell cell) => { return cell?._GetValueAsString(true) == valueAsString; };
                    }
                }
                return rows.Where(a =>
                {
                    if (a == null)
                        return false;
                    foreach (var rk in rowKeys)
                        if (!rk.IsValueMatch(a.GetCell(rk.X - 1)))
                            return false;
                    return true;
                });
            }

            ///// <summary>
            ///// (!)Uses the cache only if the cache is initialized by UseCachedDataRows=true.
            ///// Otherwise it re-reads the file every call which can be slow.
            ///// </summary>
            ///// <param name="rowKeys"></param>
            ///// <returns></returns>
            //public IEnumerable<IRow> FindRows(params NamedValue[] rowKeys)
            //{
            //    List<IRow> dataRows = cachedDataRows != null ? cachedDataRows : getDataRows().ToList();
            //    return FindRows(dataRows, rowKeys);
            //}

            /// <summary>
            /// (!)Re-reads the sheet on every call which can be slow.
            /// </summary>
            /// <param name="rowKeys"></param>
            /// <returns></returns>
            public IEnumerable<IRow> FindDataRows(params NamedValue[] rowKeys)
            {
                return FindRows(Sheet._GetRows(RowScope.WithCells).Skip(1), rowKeys);
            }

            public IEnumerable<IRow> GetDataRows(RowScope rowScope)
            {
                return Sheet._GetRows(rowScope).Skip(1);
            }

            public class NamedValue
            {
                public Column Column { get; internal set; }
                public object Value { get; internal set; }
                public int X { get { return Column.X; } }
                public Func<ICell, bool> IsValueMatch = null;

                public NamedValue(Column column, object value, Func<ICell, bool> isValueMatch = null)
                {
                    if (column.X <= 0)
                        throw new Exception("Column has no X set.");
                    Column = column;
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, string header, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(header);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, Func<string, bool> isHeaderMatch, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(isHeaderMatch);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, Regex headerRegex, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(headerRegex);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }
            }

            public IRow AppendRow<T>(IEnumerable<T> values)
            {
                IRow r = Sheet._AppendRow(values);
                setColumnStyles(r);

                //if (cachedDataRows != null)
                //    cachedDataRows.Add(r);

                return r;
            }

            void setColumnStyles(IRow row)
            {
                foreach (Column c in Columns)
                    row._GetCell(c.X, true).CellStyle = c.DataStyle;
            }

            public IRow AppendRow(params string[] values)
            {
                return AppendRow((IEnumerable<string>)values);
            }

            public IRow AppendRow(IEnumerable<NamedValue> namedValues)
            {
                IRow r = writeRow(Sheet._GetLastRow(LastRowCondition.HasCells, false) + 1, namedValues);
                return r;
            }

            IRow writeRow(int y, IEnumerable<NamedValue> namedValues)
            {
                IRow r = Sheet.GetRow(y - 1);
                if (r == null)
                {
                    r = Sheet.CreateRow(y - 1);
                    setColumnStyles(r);

                    //if (cachedDataRows != null)
                    //    cachedDataRows.Insert(r.RowNum, r);
                }
                foreach (var nv in namedValues)
                {
                    var c = r._GetCell(nv.X, true);
                    c._SetValue(nv.Value);
                }
                return r;
            }

            public IRow AppendRow(params NamedValue[] values)
            {
                return AppendRow((IEnumerable<NamedValue>)values);
            }

            public IRow InsertRow<T>(int y, IEnumerable<T> values = null)
            {
                IRow r = Sheet._InsertRow(y, values);
                setColumnStyles(r);

                //if (cachedDataRows != null)
                //    cachedDataRows.Insert(r.RowNum, r);

                return r;
            }

            public IRow InsertRow(int y, params string[] values)
            {
                return InsertRow(y, (IEnumerable<string>)values);
            }

            public IRow InsertRow(int y, IEnumerable<NamedValue> namedValues)
            {
                int lastRowY = Sheet._GetLastRow(LastRowCondition.HasCells, false);
                if (y <= lastRowY)
                    Sheet.ShiftRows(y - 1, lastRowY - 1, 1);
                IRow r = writeRow(y, namedValues);
                return r;
            }

            public IRow InsertRow(int y, params NamedValue[] values)
            {
                return InsertRow(y, (IEnumerable<NamedValue>)values);
            }

            public IRow WriteRow<T>(int y, IEnumerable<T> values = null)
            {
                IRow r = Sheet._WriteRow(y, values);
                return r;
            }

            public IRow WriteRow(int y, params string[] values)
            {
                return WriteRow(y, (IEnumerable<string>)values);
            }

            public IRow WriteRow(int y, IEnumerable<NamedValue> namedValues)
            {
                IRow r = writeRow(y, namedValues);
                return r;
            }

            public IRow WriteRow(int y, params NamedValue[] values)
            {
                return WriteRow(y, (IEnumerable<NamedValue>)values);
            }

            public IRow RemoveRow(int y, bool shiftRemainingRows)
            {
                return Sheet._RemoveRow(y, shiftRemainingRows);

                //if (cachedDataRows != null && r != null)
                //    cachedDataRows.Remove(r);
            }

            public ICell GetCell(IRow row, string header, bool create)
            {
                return row._GetCell(GetColumn(header).X, create);
            }

            public ICell GetCell(IRow row, Column column, bool create)
            {
                return row._GetCell(column.X, create);
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