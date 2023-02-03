//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using System.Collections.ObjectModel;
using static Cliver.Excel.Table;

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
                headers = new ReadOnlyCollection<string>(headersRow.GetCells(true).Select(a => a.GetValueAsString()).ToList());
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
                    Excel.GetRow(1, true).Write(headers);
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
                        if (a.GetCell(rk.ColumnX - 1)?.GetValue() != rk.Value)
                            return false;
                    return true;
                });
            }

            public IEnumerable<IRow> FindRows(params NamedValue[] rowKeys)
            {
                return FindRows(/*cachedRows*/Excel.GetRowsInRange(false, 2), rowKeys);
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

            public IRow AppendRow(IEnumerable<object> values)
            {
                IRow r = Excel.AppendRow(values);
                //cachedRows.Add(r);
                return r;
            }

            public IRow AppendRow(params object[] values)
            {
                return AppendRow((IEnumerable<object>)values);
            }

            public IRow AppendRow(IEnumerable<NamedValue> namedValues)
            {
                int y = Excel.Sheet.LastRowNum + 2;
                IRow r = writeRow(y, namedValues);
                //cachedRows.Add(r);
                return r;
            }

            IRow writeRow(int y, IEnumerable<NamedValue> namedValues)
            {
                var r = Excel.GetRow(y, true);
                foreach (var nv in namedValues)
                {
                    var c = r.GetCell(nv.ColumnX, true);
                    c.SetValue(nv.Value);
                }
                return r;
            }

            public IRow AppendRow(params NamedValue[] values)
            {
                return AppendRow((IEnumerable<NamedValue>)values);
            }

            public IRow InsertRow(int y, IEnumerable<object> values = null)
            {
                IRow r = Excel.InsertRow(y, values);
                //cachedRows.Insert(r.RowNum, r);
                return r;
            }

            public IRow InsertRow(int y, params object[] values)
            {
                return InsertRow(y, (IEnumerable<object>)values);
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

            public IRow WriteRow(int y, IEnumerable<object> values = null)
            {
                IRow r = Excel.WriteRow(y, values);
                //cachedRows[r.RowNum] = r;
                return r;
            }

            public IRow WriteRow(int y, params object[] values)
            {
                return WriteRow(y, (IEnumerable<object>)values);
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
                return row.GetCell(GetHeaderX(header), create);
            }

            public void SetStyles(IRow row, params ICellStyle[] styles)
            {
                row.SetStyles(1, styles);
            }

            public void SetStyles(IRow row, IEnumerable<ICellStyle> styles)
            {
                SetStyles(row, styles.ToArray());
            }
        }
    }
}