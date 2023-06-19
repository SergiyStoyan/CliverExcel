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
            public ReadOnlyCollection<Column> Columns { get; private set; }

            public void SetColumns(SetColumnMode setColumnMode, params string[] headers)
            {
                SetColumns(setColumnMode, (IEnumerable<string>)headers);
            }

            public void SetColumns(SetColumnMode setColumnMode, IEnumerable<string> headers)
            {
                SetColumns(setColumnMode, headers.Select(a => new Column(a)));
            }

            public void SetColumns(SetColumnMode setColumnMode, params Column[] columns)
            {
                SetColumns(setColumnMode, (IEnumerable<Column>)columns);
            }

            public enum SetColumnMode
            {
                OverrideAll,
                KeepExisting,
                ThrowExceptionIfDiffer
            }
            public void SetColumns(SetColumnMode setColumnMode, IEnumerable<Column> columns)
            {
                var duplicates = columns.GroupBy(a => a.Header).Where(a => a.Count() > 1).Select(a => "'" + a.Key + "'").ToList();
                if (duplicates.Count > 0)
                    throw new Exception("Columns duplicated: " + string.Join(", ", duplicates));

                switch (setColumnMode)
                {
                    case SetColumnMode.OverrideAll:
                        Columns = new ReadOnlyCollection<Column>(columns?.ToList());
                        break;

                    case SetColumnMode.KeepExisting:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> ccs = Columns.ToList();
                            for (int i = cs.Count - 1; i >= 0; i--)
                            {
                                Column c = cs[i];
                                var cc = Columns.FirstOrDefault(a => c.IsHeaderMatch(a.Header));
                                if (cc != null)
                                {
                                    ccs.Remove(cc);
                                    ccs.Insert(cc.X - 1, c);
                                    cs.RemoveAt(i);
                                }
                            }
                            Columns = new ReadOnlyCollection<Column>(ccs.Concat(cs).ToList());
                        }
                        break;

                    case SetColumnMode.ThrowExceptionIfDiffer:
                        {
                            List<Column> cs = columns.ToList();
                            if (cs.Count > Columns.Count)
                                throw new Exception("The number of existing columns " + Columns.Count + " < the number of new columns " + cs.Count);
                            int i = 0;
                            foreach (Column c in columns)
                            {
                                Column cc = Columns[i++];
                                if (!c.IsHeaderMatch(cc.Header))
                                    throw new Exception("Existing column[x=" + cc.X + "] '" + cc.Header + "' differs from the new one '" + c.Header + "'");
                            }
                            Columns = new ReadOnlyCollection<Column>(columns?.ToList());
                        }
                        break;

                    default:
                        throw new Exception("Unknown case: " + setColumnMode);
                }

                Columns?.Select((a, i) => (column: a, x: i + 1)).ForEach(a =>
                {
                    a.column.X = a.x;
                    a.column.Table = this;
                });
                Sheet._GetRow(1, true)._Write(Columns?.Select((a, i) => a.Header));
                var r2 = Sheet._GetRow(2, false);
                if (r2 != null)
                    Columns.ForEach(a => a.SetDataStyle(r2._GetCell(a.X, false)?.CellStyle, false));
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

            public void InsertColumn(int x, Column column)
            {
                Sheet._ShiftColumnsRight(x, 1);
                column.X = x;
                Sheet._GetCell(1, x, true)._SetValue(column.Header);
                var cs = Columns.ToList();
                cs.Insert(column.X - 1, column);
                Columns = new ReadOnlyCollection<Column>(cs);
            }

            public void RemoveColumn(Column column)
            {
                Sheet._ShiftColumnsLeft(column.X, 1);
                var cs = Columns.ToList();
                cs.RemoveAt(column.X - 1);
                Columns = new ReadOnlyCollection<Column>(cs);
            }

            public class Column
            {
                public readonly string Header;
                public int X { get; internal set; } = -1;

                public ICellStyle DataStyle { get; private set; } = null;
                public void SetDataStyle(ICellStyle style, bool updateExistingCells)
                {
                    DataStyle = style;
                    if (updateExistingCells)
                        foreach (ICell c in GetDataCells(RowScope.WithCells))
                            c.CellStyle = DataStyle;
                }

                public Table Table { get; internal set; } = null;

                public Func<string, bool> IsHeaderMatch { get; internal set; }

                /// <summary>
                /// (!)Until a new column is passed to Excel.Table.Columns, it remains non-initialized.
                /// </summary>
                /// <param name="header"></param>
                /// <param name="style"></param>
                public Column(string header, ICellStyle dataStyle = null, Func<string, bool> isHeaderMatch = null)
                {
                    if (header == null)
                        throw new ArgumentNullException("header");
                    Header = header;
                    SetDataStyle(dataStyle, false);
                    IsHeaderMatch = isHeaderMatch != null ? isHeaderMatch : (h) => { return h == Header; };
                }

                public ICell GetCell(int y, bool create)
                {
                    return Table.GetCell(y, this, create);
                }

                public IEnumerable<ICell> GetDataCells(RowScope rowScope)
                {
                    return Table.Sheet._GetRowsInRange(rowScope, 2).Select(a => a?.GetCell(X));
                }
            }
        }
    }
}