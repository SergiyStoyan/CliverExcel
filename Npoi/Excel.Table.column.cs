﻿//********************************************************************************************
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
                /// <summary>
                /// The listed columns override the header row content.
                /// </summary>
                OverrideAll,
                /// <summary>
                /// The listed columns, that are not found in the header row in any order, are added to the right.
                /// </summary>
                ExpandOldHeaders,
                /// <summary>
                /// The listed columns all must be found in the header row in any order.
                /// </summary>
                FindHeaders,
                /// <summary>
                /// The listed columns all must be found in the header row in the listed order.
                /// </summary>
                FindHeadersOrdered,
                /// <summary>
                /// If the header row is empty then the listed columns are set. Otherwise they all must be found in the header row in any order.
                /// </summary>
                CreateOrFindHeaders,
                /// <summary>
                /// If the header row is empty then the listed columns are set. Otherwise they all must be found in the header row in the listed order.
                /// </summary>
                CreateOrFindHeadersOrdered,
            }
            /// <summary>
            /// (!)NULLs among input columns are allowed. They make gaps between columns but they are not listed in Columns.
            /// </summary>
            /// <param name="setColumnMode"></param>
            /// <param name="columns"></param>
            /// <exception cref="Exception"></exception>
            public void SetColumns(SetColumnMode setColumnMode, IEnumerable<Column> columns)
            {
                switch (setColumnMode)
                {
                    case SetColumnMode.OverrideAll:
                        columns = columns
                            .Select((a, i) => (column: a, x: i + 1))
                            .Where(a =>
                            {
                                if (a.column == null)
                                    return false;
                                a.column.X = a.x;
                                return true;
                            })
                            .Select(a => a.column);
                        Columns = new ReadOnlyCollection<Column>(columns.ToList());
                        break;

                    case SetColumnMode.ExpandOldHeaders:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> c0s = Columns.ToList();
                            int lastX = c0s[c0s.Count - 1].X;
                            for (int i = cs.Count - 1; i >= 0; i--)
                            {
                                Column c = cs[i];
                                if (c == null)
                                {
                                    cs.RemoveAt(i);
                                    continue;
                                }
                                for (int j = c0s.Count - 1; j >= 0; j--)
                                {
                                    Column c0 = c0s[j];
                                    if (c.IsHeaderMatch(c0.Header))
                                    {
                                        c.X = c0.X;
                                        c0s.RemoveAt(j);
                                        c0s.Insert(j, c);
                                        cs.RemoveAt(i);
                                        break;
                                    }
                                }
                            }
                            cs = cs.Select((a, i) => (column: a, x: c0s.Count + i))
                                .Where(a =>
                                {
                                    if (a.column == null)
                                        return false;
                                    a.column.X = a.x;
                                    return true;
                                })
                                .Select(a => a.column)
                                .ToList();
                            Columns = new ReadOnlyCollection<Column>(c0s.Concat(cs).ToList());
                        }
                        break;

                    case SetColumnMode.FindHeaders:
                        {
                            List<Column> cs = columns.Where(a => a != null).ToList();
                            if (cs.Count > Columns.Count)
                                throw new Exception("The number of existing columns " + Columns.Count + " < the number of new columns " + cs.Count);
                            List<Column> c0s = Columns.ToList();
                            for (int i = cs.Count - 1; i >= 0; i--)
                            {
                                Column c = cs[i];
                                int j = c0s.Count - 1;
                                for (; j >= 0; j--)
                                {
                                    Column c0 = c0s[j];
                                    if (c.IsHeaderMatch(c0.Header))
                                    {
                                        c.X = c0.X;
                                        c0s.RemoveAt(j);
                                        c0s.Insert(j, c);
                                        break;
                                    }
                                }
                                if (j < 0)
                                    throw new Exception("Column '" + c.Header + "' does not exist in the table.");
                            }
                            Columns = new ReadOnlyCollection<Column>(c0s);
                        }
                        break;

                    case SetColumnMode.FindHeadersOrdered:
                        {
                            List<Column> cs = columns.ToList();
                            int notEmptyCount = cs.Where(a => a != null).Count();
                            if (notEmptyCount > Columns.Count)
                                throw new Exception("The number of existing columns " + Columns.Count + " < the number of new columns " + notEmptyCount);
                            List<Column> c0s = Columns.ToList();
                            int emptyCount = 0;
                            for (int i = 0; i < cs.Count; i++)
                            {
                                Column c = cs[i];
                                if (c == null)
                                {
                                    emptyCount++;
                                    if (c0s[i].X < i + emptyCount)
                                        throw new Exception("NULL column[position=" + i + "] does not exist in the table.");
                                    continue;
                                }
                                Column c0 = c0s[i + emptyCount];
                                if (!c.IsHeaderMatch(c0.Header))
                                    throw new Exception("Existing column[x=" + c0.X + "] '" + c0.Header + "' differs from the new one '" + c.Header + "'");
                                c.X = c0.X;
                                c0s.RemoveAt(i);
                                c0s.Insert(i, c);
                            }
                            Columns = new ReadOnlyCollection<Column>(c0s);
                        }
                        break;

                    case SetColumnMode.CreateOrFindHeaders:
                        if (Columns.Any())
                            goto case SetColumnMode.FindHeaders;
                        goto case SetColumnMode.OverrideAll;

                    case SetColumnMode.CreateOrFindHeadersOrdered:
                        if (Columns.Any())
                            goto case SetColumnMode.FindHeadersOrdered;
                        goto case SetColumnMode.OverrideAll;

                    default:
                        throw new Exception("Unknown case: " + setColumnMode);
                }

                Columns.ForEach(a => { a.Table = this; });

                for (int i = 0; i < Columns.Count; i++)
                {
                    Column c = Columns[i];
                    for (int j = i + 1; j < Columns.Count; j++)
                    {
                        Column cj = Columns[j];
                        if (cj.X == c.X)
                            throw new Exception("Columns have the same X: '" + c.Header + "'[x=" + c.X + "] == '" + cj.Header + "'[x=" + cj.X + "]");
                        if (cj.IsHeaderMatch(c.Header))
                            throw new Exception("Columns are equal by IsHeaderMatch(): '" + c.Header + "'[x=" + c.X + "] == '" + cj.Header + "'[x=" + cj.X + "]");
                    }
                }

                WriteRow(1, Columns.Select(a => new Cell(a, a.Header)));

                var r2 = Sheet._GetRow(2, false);
                if (r2 != null)
                    Columns.ForEach(a => a.SetDataStyle(r2._GetCell(a.X, false)?.CellStyle, false));
            }

            public Column GetColumn(string header, bool exceptionIfNotFound = true)
            {
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
                if (column == null)
                    return;
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
                    if (string.IsNullOrWhiteSpace(header))
                        throw new Exception("Header cannot be empty space.");
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