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
            public Table(Excel excel)
            {
                Excel = excel;
                IRow headersRow = excel.GetRow(1, true);
                headers = new ReadOnlyCollection<string>(headersRow._GetCells(true).Select(a => a._GetValueAsString()).ToList());
            }

            public Table(Excel excel, params string[] headers)
            {
                Excel = excel;
                Headers = new ReadOnlyCollection<string>(headers.ToList());
            }

            readonly public Excel Excel;

            public ReadOnlyCollection<string> Headers
            {
                get { return headers; }
                set
                {
                    headers = value;
                    Excel.GetRow(1, true)._Write(headers);
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

            ///// <summary>
            ///// Must be called when the sheet was edited outside this class.
            ///// </summary>
            //public void ReloadCachedRows()
            //{
            //    _cachedRows = new List<IRow>();
            //    foreach (IRow r in Excel.GetRows(false, false))
            //    {
            //        while (_cachedRows.Count < r.RowNum)
            //            _cachedRows.Add(null);
            //        _cachedRows.Add(r);
            //    }
            //}

            //List<IRow> cachedRows
            //{
            //    get
            //    {
            //        if (_cachedRows == null)
            //            ReloadCachedRows();
            //        return _cachedRows;
            //    }
            //}
            //List<IRow> _cachedRows;

            static public IEnumerable<IRow> FindRows(IEnumerable<IRow> rows, params NamedValue[] rowKeys)
            {
                return rows.Where(a =>
                {
                    foreach (var rk in rowKeys)
                        if (a.GetCell(rk.ColumnX - 1)?._GetValue() != rk.Value)
                            return false;
                    return true;
                });
            }

            public IEnumerable<IRow> FindRows(params NamedValue[] rowKeys)
            {
                return FindRows(/*cachedRows*/Excel.GetRowsInRange(RowScope.OnlyExisting, 2), rowKeys);
            }

            public NamedValue NewNamedValue(string header, object value)
            {
                return new NamedValue(header, value, GetHeaderX(header));
            }

            public class NamedValue
            {
                public string Header { get; internal set; }
                public object Value { get; internal set; }
                public int ColumnX { get; internal set; }

                internal NamedValue(string header, object value, int columnX)
                {
                    Header = header;
                    Value = value;
                    ColumnX = columnX;
                }
            }

            public IRow AppendRow<T>(IEnumerable<T> values)
            {
                IRow r = Excel.AppendRow(values);
                //cachedRows.Add(r);
                return r;
            }

            public IRow AppendRow(params string[] values)
            {
                return AppendRow((IEnumerable<string>)values);
            }

            public IRow AppendRow(IEnumerable<NamedValue> namedValues)
            {
                int y0 = Excel.Sheet.LastRowNum;//(!)it is 0 when no row or 1 row
                int y = y0 + (y0 == 0 && Excel.Sheet.GetRow(y0) == null ? 1 : 2);
                IRow r = writeRow(y, namedValues);
                //cachedRows.Add(r);
                return r;
            }

            IRow writeRow(int y, IEnumerable<NamedValue> namedValues)
            {
                var r = Excel.GetRow(y, true);
                foreach (var nv in namedValues)
                {
                    var c = r._GetCell(nv.ColumnX, true);
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
                IRow r = Excel.InsertRow(y, values);
                //cachedRows.Insert(r.RowNum, r);
                return r;
            }

            public IRow InsertRow(int y, params string[] values)
            {
                return InsertRow(y, (IEnumerable<string>)values);
            }

            public IRow InsertRow(int y, IEnumerable<NamedValue> namedValues)
            {
                if (y <= Excel.Sheet.LastRowNum + 1)
                    Excel.Sheet.ShiftRows(y - 1, Excel.Sheet.LastRowNum, 1);
                IRow r = writeRow(y, namedValues);
                //cachedRows.Insert(r.RowNum, r);
                return r;
            }

            public IRow InsertRow(int y, params NamedValue[] values)
            {
                return InsertRow(y, (IEnumerable<NamedValue>)values);
            }

            public IRow WriteRow<T>(int y, IEnumerable<T> values = null)
            {
                IRow r = Excel.WriteRow(y, values);
                //cachedRows[r.RowNum] = r;
                return r;
            }

            public IRow WriteRow(int y, params string[] values)
            {
                return WriteRow(y, (IEnumerable<string>)values);
            }

            public IRow WriteRow(int y, IEnumerable<NamedValue> namedValues)
            {
                IRow r = writeRow(y, namedValues);
                //cachedRows[r.RowNum] = r;
                return r;
            }

            public IRow WriteRow(int y, params NamedValue[] values)
            {
                return WriteRow(y, (IEnumerable<NamedValue>)values);
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
        }
    }
}