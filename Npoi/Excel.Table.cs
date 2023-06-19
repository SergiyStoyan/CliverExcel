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
        public partial class Table
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
                ////allow 1 empty header (as unique)
                //var hs = headersRow._GetCells(true).Select(a => a._GetValueAsString()).TakeWhile(a => !string.IsNullOrWhiteSpace(a) || ++emptyCount < 2).ToList();
                //allow 1 empty-header column (as unique) and ignore the rest empty-header columns
                var hs = headersRow._GetCells(true).Select(a => a._GetValueAsString()).Where(a => !string.IsNullOrWhiteSpace(a) || ++emptyCount < 2).ToList();
                SetColumns(SetColumnMode.OverrideAll, hs);
            }

            public Table(Excel excel, SetColumnMode setColumnMode, params string[] headers) : this(excel, setColumnMode, (IEnumerable<string>)headers) { }

            public Table(Excel excel, SetColumnMode setColumnMode, IEnumerable<string> headers) : this(excel)
            {
                SetColumns(setColumnMode, headers);
            }

            public Table(Excel excel, SetColumnMode setColumnMode, params Column[] columns) : this(excel, setColumnMode, (IEnumerable<Column>)columns) { }

            public Table(Excel excel, SetColumnMode setColumnMode, IEnumerable<Column> columns) : this(excel)
            {
                SetColumns(setColumnMode, columns);
            }

            readonly public ISheet Sheet;
            readonly public Excel Excel;

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

            public IRow InsertRow(int y, params NamedValue[] values)
            {
                return InsertRow(y, (IEnumerable<NamedValue>)values);
            }

            public IRow InsertRow(int y, IEnumerable<NamedValue> namedValues)
            {
                int lastRowY = Sheet._GetLastRow(LastRowCondition.HasCells, false);
                if (y <= lastRowY)
                    Sheet.ShiftRows(y - 1, lastRowY - 1, 1);
                IRow r = writeRow(y, namedValues);
                return r;
            }

            //public IRow InsertFullRow<T>(int y, IEnumerable<T> values = null)
            //{if()
            //    IRow r = Sheet._InsertRow(y, values);
            //    setColumnStyles(r);

            //    //if (cachedDataRows != null)
            //    //    cachedDataRows.Insert(r.RowNum, r);

            //    return r;
            //}

            //public IRow InsertFullRow(int y, params string[] values)
            //{
            //    return InsertFullRow(y, (IEnumerable<string>)values);
            //}

            //public IRow InsertFullRow(int y, params NamedValue[] values)
            //{
            //    return InsertFullRow(y, (IEnumerable<NamedValue>)values);
            //}

            //public IRow InsertFullRow(int y, IEnumerable<NamedValue> namedValues)
            //{
            //    return InsertRow(y, (IEnumerable<NamedValue>)values);
            //}

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

            /// <summary>
            /// (!)Seeks the column each call.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="header"></param>
            /// <param name="create"></param>
            /// <returns></returns>
            public ICell GetCell(IRow row, string header, bool create)
            {
                return row._GetCell(GetColumn(header).X, create);
            }

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