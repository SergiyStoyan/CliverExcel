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
                headers = new ReadOnlyCollection<string>(headersRow._GetCells(true).Select(a => a._GetValueAsString()).ToList());
            }

            public Table(Excel excel, params string[] headers)
            {
                Excel = excel;
                Sheet = Excel.Sheet;
                Headers = new ReadOnlyCollection<string>(headers);
            }

            readonly public ISheet Sheet;
            readonly public Excel Excel;

            public ReadOnlyCollection<string> Headers
            {
                get { return headers; }
                set
                {
                    headers = value;
                    Sheet._GetRow(1, true)._Write(headers);
                }
            }
            ReadOnlyCollection<string> headers;

            /// <summary>
            /// 
            /// </summary>
            /// <param name="header"></param>
            /// <returns>1-based</returns>
            /// <exception cref="Exception"></exception>
            public int GetHeaderX(string header)
            {
                int i = headers.IndexOf(header);
                if (i < 0)
                {
                    //headers.Add(nv.Header);
                    //i = headers.Count - 1;
                    throw new Exception2("There is no header '" + header + "'");
                }
                return i + 1;
            }

            public void SetColumnStyle(string header, ICellStyle style)
            {
                int x = GetHeaderX(header);
                xs2columnStyle[x] = new ColumnStyle { X = x, Style = style };
            }

            internal class ColumnStyle
            {
                internal ICellStyle Style;
                internal int X;
            }
            Dictionary<int, ColumnStyle> xs2columnStyle = new Dictionary<int, ColumnStyle>();

            public ICellStyle GetColumnStyle(string header)
            {
                int x = GetHeaderX(header);
                xs2columnStyle.TryGetValue(x, out ColumnStyle cs);
                return cs?.Style;
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
                foreach (var rk in rowKeys)
                    rk.ValueAsString = rk.Value?.ToString();

                return rows.Where(a =>
                {
                    if (a == null)
                        return false;
                    foreach (var rk in rowKeys)
                        if (a.GetCell(rk.X - 1)?._GetValue().ToString() != rk.ValueAsString)
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

            public NamedValue NewNamedValue(string header, object value)
            {
                return new NamedValue(header, value, GetHeaderX(header));
            }

            public class NamedValue
            {
                public string Header { get; internal set; }
                public object Value { get; internal set; }
                public int X { get; internal set; }

                internal string ValueAsString;

                internal NamedValue(string header, object value, int columnX)
                {
                    Header = header;
                    Value = value;
                    X = columnX;
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
                foreach (ColumnStyle cs in xs2columnStyle.Values)
                    if (cs != null)
                        row._GetCell(cs.X, true).CellStyle = cs.Style;
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
                return row._GetCell(GetHeaderX(header), create);
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