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
            /// <summary>
            /// All the columns in the Table. In this list they are always registered/initialized.
            /// </summary>
            public ReadOnlyCollection<Column> Columns { get; private set; }

            //public void SetColumns(SetColumnMode setColumnMode, params string[] headers)
            //{
            //    SetColumns(setColumnMode, (IEnumerable<string>)headers);
            //}

            //public void SetColumns(SetColumnMode setColumnMode, IEnumerable<string> headers)
            //{
            //    SetColumns(setColumnMode, headers.Select(a => new Column(a)));
            //}

            /// <summary>
            /// Registers/initializes the listed columns. It is a necessary call in the beginning of using Excel.Table.
            /// (!)NULLs among input columns are allowed. They make gaps between columns but they do not go to Columns.
            /// </summary>
            /// <param name="setColumnMode"></param>
            /// <param name="columns"></param>
            public void SetColumns(SetColumnMode setColumnMode, params Column[] columns)
            {
                SetColumns(setColumnMode, (IEnumerable<Column>)columns);
            }

            public enum SetColumnMode
            {
                /// <summary>
                /// The listed columns override the header row content.
                /// </summary>
                Override,
                /// <summary>
                /// The listed columns, that are not found in the header row in any order, are added to the right.
                /// </summary>
                FindOrAppend,
                /// <summary>
                /// The listed columns must exist in the header row in any order.
                /// </summary>
                Find,
                /// <summary>
                /// The listed columns must exist in the header row in the listed order.
                /// </summary>
                FindOrdered,
                /// <summary>
                /// The listed columns must exist in the header row in any order, or the header row must be empty in which case the listed columns are created.
                /// </summary>
                FindOrCreate,
                /// <summary>
                /// The listed columns must exist in the header row in the listed order, or the header row must be empty in which case the listed columns are created.
                /// </summary>
                FindOrderedOrCreate,
            }

            /// <summary>
            /// Registers/initializes the listed columns. It is a necessary call before using Excel.Table, which can be made in a derived Table's constructor.
            /// (!)NULLs among input columns are allowed. They make gaps between columns but they do not go into Columns.
            /// </summary>
            /// <param name="setColumnMode"></param>
            /// <param name="columns"></param>
            /// <exception cref="Exception"></exception>
            public void SetColumns(SetColumnMode setColumnMode, IEnumerable<Column> columns)
            {
                switch (setColumnMode)
                {
                    case SetColumnMode.Override:
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

                    case SetColumnMode.FindOrAppend:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> c0s = Columns.ToList();
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
                                    if (c.Header == c0.Header)
                                    {
                                        c.X = c0.X;
                                        c0s.RemoveAt(j);
                                        c0s.Insert(j, c);
                                        cs.RemoveAt(i);
                                        break;
                                    }
                                }
                            }
                            cs = cs.Select((a, i) => (column: a, x: c0s.Count + i + 1))
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

                    case SetColumnMode.Find:
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
                                    if (c.Header == c0.Header)
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

                    case SetColumnMode.FindOrdered:
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
                                if (c.Header != c0.Header)
                                    throw new Exception("Existing column[x=" + c0.X + "] '" + c0.Header + "' differs from the new one '" + c.Header + "'");
                                c.X = c0.X;
                                c0s.RemoveAt(i);
                                c0s.Insert(i, c);
                            }
                            Columns = new ReadOnlyCollection<Column>(c0s);
                        }
                        break;

                    case SetColumnMode.FindOrCreate:
                        if (Columns.Any())
                            goto case SetColumnMode.Find;
                        goto case SetColumnMode.Override;

                    case SetColumnMode.FindOrderedOrCreate:
                        if (Columns.Any())
                            goto case SetColumnMode.FindOrdered;
                        goto case SetColumnMode.Override;

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
                        if (cj.Header == c.Header)
                            throw new Exception("Columns are equal by headers: '" + c.Header + "'[x=" + c.X + "] == '" + cj.Header + "'[x=" + cj.X + "]");
                    }
                }

                WriteRow(1, Columns.Select(a => new Cell(a, a.Header)));

                {
                    var r2 = Sheet._GetRow(2, false);
                    bool r2created = false;
                    if (r2 == null)
                    {
                        r2 = Sheet._GetRow(2, true);
                        r2created = true;
                    }
                    Columns.Where(a => a.DataStyle == null).ForEach(a =>
                    {
                        ICell c = r2._GetCell(a.X, false);
                        bool ccreated = false;
                        if (c == null)
                        {
                            c = r2._GetCell(a.X, true);
                            ccreated = true;
                        }
                        a.DataStyle = c.CellStyle;
                        if (ccreated)
                            c._Remove();
                    });
                    Columns.Where(a => a.DataType == null).ForEach(a =>
                    {
                        ICell c = r2._GetCell(a.X, false);
                        bool ccreated = false;
                        if (c == null)
                        {
                            c = r2._GetCell(a.X, true);
                            ccreated = true;
                        }
                        a.DataType = c.CellType;
                        if (ccreated)
                            c._Remove();
                    });
                    if (r2created)
                        r2._Remove();
                }
            }

            /// <summary>
            /// Find a registered column matched by header.
            /// </summary>
            /// <param name="header"></param>
            /// <param name="exceptionIfNotFound"></param>
            /// <returns></returns>
            public Column GetColumn(string header, bool exceptionIfNotFound = true)
            {
                return GetColumn((v) => { return v == header; }, exceptionIfNotFound);
            }

            /// <summary>
            /// Find a registered column matched by the input column's header.
            /// </summary>
            /// <param name="column"></param>
            /// <param name="exceptionIfNotFound"></param>
            /// <returns></returns>
            public Column GetColumn(Column column, bool exceptionIfNotFound = true)
            {
                return GetColumn((v) => { return column.Header == v; }, exceptionIfNotFound);
            }

            /// <summary>
            /// Find a registered column matched by headerMatchRegex.
            /// </summary>
            /// <param name="headerMatchRegex"></param>
            /// <param name="exceptionIfNotFound"></param>
            /// <returns></returns>
            public Column GetColumn(Regex headerMatchRegex, bool exceptionIfNotFound = true)
            {
                return GetColumn((v) => { return headerMatchRegex.IsMatch(v); }, exceptionIfNotFound);
            }

            /// <summary>
            /// Find a registered column matched by isHeaderMatch.
            /// </summary>
            public Column GetColumn(Func<string, bool> isHeaderMatch, bool exceptionIfNotFound = true)
            {
                var c = Columns.FirstOrDefault(a => isHeaderMatch(a.Header));
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
                SetColumns(SetColumnMode.FindOrdered, cs);
            }

            public void RemoveColumn(Column column)
            {
                if (column.Table == null)
                    throw new Exception("Column is not initialized: Table is not set.");
                Sheet._ShiftColumnsLeft(column.X, 1);
                var cs = Columns.ToList();
                cs.RemoveAt(column.X - 1);
                SetColumns(SetColumnMode.FindOrdered, cs);
            }

            public class Column
            {
                public readonly string Header;
                public int X { get; internal set; } = -1;

                /// <summary>
                /// (!)Unregistered style will be registered when setting.
                /// </summary>
                public ICellStyle DataStyle
                {
                    get
                    {
                        return dataStyle;
                    }
                    set
                    {
                        if (value == null)
                            return;
                        if (value.Index < 0)
                            value = Table.Excel.GetRegisteredStyle(value);
                        dataStyle = value;
                        //Table?.Sheet.SetDefaultColumnStyle(X - 1, dataStyle);
                    }
                }
                ICellStyle dataStyle = null;
                public void ApplyDataStyle(ICellStyle style = null)
                {
                    if (style != null)
                        DataStyle = style;
                    foreach (ICell c in GetDataCells(RowScope.WithCells))
                        c.CellStyle = DataStyle;
                }

                public CellType? DataType { get; set; } = null;
                public void ApplyDataType(CellType? dataType = null)
                {
                    if (dataType != null)
                        DataType = dataType.Value;
                    if (DataType != null)
                        foreach (ICell c in GetDataCells(RowScope.WithCells))
                            c.SetCellType(DataType.Value);
                }

                public Table Table { get; internal set; } = null;

                /// <summary>
                /// (!)Until a created column is registered in Excel.Table.Columns, it is not initialized and cannot be used in most methods.
                /// </summary>
                /// <param name="header"></param>
                /// <param name="style"></param>
                public Column(string header, ICellStyle dataStyle = null, CellType? dataType = null)
                {
                    if (string.IsNullOrWhiteSpace(header))
                        throw new Exception("Header cannot be empty or space.");
                    Header = header;
                    DataStyle = dataStyle;
                    DataType = dataType;
                }

                public ICell GetCell(int y, bool create)
                {
                    return Table.GetCell(y, this, create);
                }

                public IEnumerable<ICell> GetDataCells(RowScope rowScope)
                {
                    return Table.Sheet._GetRowsInRange(rowScope, 2).Select(a => a?.GetCell(X));
                }

                /// <summary>
                /// (!)Unregistered style will be registered.
                /// </summary>
                /// <param name="value"></param>
                /// <param name="style"></param>
                /// <param name="type"></param>
                /// <returns></returns>
                public Cell NewCell(object value, ICellStyle style = null, CellType? type = null)
                {
                    return new Cell(this, value, style, type);
                }

                /// <summary>
                /// (!)Unregistered style will be registered.
                /// </summary>
                /// <param name="style"></param>
                /// <returns></returns>
                public Style NewStyle(ICellStyle style)
                {
                    return new Style(this, style);
                }
            }
        }
    }
}