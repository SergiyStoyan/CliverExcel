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

            void loadColumns()
            {
                IRow headersRow = Sheet._GetRow(1, true);
                IEnumerable<Column> columns = headersRow._GetCells(CellScope.CreateIfNull).Select(a =>
                {
                    string h = a._GetValueAsString();
                    return string.IsNullOrWhiteSpace(h) ? null : new Column(h) { X = a._X() };
                });
                setColumns(columns, false);
            }

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
                ///// <summary>
                ///// The listed columns, that are not found in the header row in any order, are inserted after the found predecessor.
                ///// </summary>
                //FindOrInsert,!!!auto-inserting/removing is not appreciated because of possibly damaging formulas and mergings
                /// <summary>
                /// The listed columns must exist in the header row in any order. The absent columns are added to the right of the table.
                /// </summary>
                Find,
                /// <summary>
                /// The listed columns must exist in the header row in the listed order.
                /// </summary>
                FindOrdered,
                ///// <summary>
                ///// The listed columns must exist in the header row in the listed order. The absent columns are created in their listed position.
                ///// </summary>
                //FindOrderedOrInsert,!!!auto-inserting/removing is not appreciated because of possibly damaging formulas and mergings
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
                        setColumns(columns, true);
                        break;

                    case SetColumnMode.FindOrAppend:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> c0s = Columns.ToList();
                            {//restore NULLs
                                for (int i = c0s.Count - 2; i >= 0; i--)
                                {
                                    Column c0 = c0s[i];
                                    for (int null0Count = c0s[i + 1].X - c0.X - 1; null0Count > 0; null0Count--)
                                        c0s.Insert(i + 1, null);
                                }
                            }
                            int lastMatchI = -1;
                            for (int i = cs.Count - 1; i >= 0; i--)
                            {
                                Column c = cs[i];
                                if (c == null)
                                {
                                    if (lastMatchI >= 0)
                                        cs.RemoveAt(i);
                                    continue;
                                }
                                for (int j = c0s.Count - 1; j >= 0; j--)
                                {
                                    Column c0 = c0s[j];
                                    if (c0 == null)
                                        continue;
                                    if (c.Header == c0.Header)
                                    {
                                        c0s.RemoveAt(j);
                                        c0s.Insert(j, c);
                                        cs.RemoveAt(i);
                                        if (lastMatchI < 0)
                                            lastMatchI = i;
                                        break;
                                    }
                                }
                            }

                            setColumns(c0s.Concat(cs), true);
                        }
                        break;

                    case SetColumnMode.Find:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> c0s = Columns.ToList();
                            int null0Count = 0;
                            for (int i0 = 0; i0 < c0s.Count; i0++)
                            {
                                Column c0 = c0s[i0];
                                int x01 = i0 - 1 < 0 ? 0 : c0s[i0 - 1].X;
                                null0Count += c0.X - x01 - 1;
                            }
                            for (int i = 0; i < cs.Count; i++)
                            {
                                Column c = cs[i];
                                if (c == null)
                                {
                                    null0Count--;
                                    if (null0Count < 0)
                                        throw new Exception("NULL column[X=" + (i + 1) + "] has no match in the table.");
                                    continue;
                                }
                                Column c0 = c0s.FirstOrDefault(b => b.Header == c.Header);
                                if (c0 == null)
                                    throw new Exception("Column[X=" + (i + 1) + "] '" + c.Header + "' has no match in the table.");
                                c0s.Remove(c0);
                            };
                        }
                        break;

                    case SetColumnMode.FindOrdered:
                        {
                            List<Column> cs = columns.ToList();
                            List<Column> c0s = Columns.ToList();
                            int i0 = 0;
                            for (int i = 0; i < cs.Count; i++)
                            {
                                Column c = cs[i];
                                if (c == null)
                                {
                                    for (; i0 < c0s.Count; i0++)
                                    {
                                        Column c0 = c0s[i0];
                                        int x01 = i0 - 1 < 0 ? 0 : c0s[i0 - 1].X;
                                        if (x01 + 1 < c0.X)
                                            break;
                                    }
                                    if (i0 >= c0s.Count)
                                        throw new Exception("NULL column[X=" + (i + 1) + "] has no match in the table.");
                                }
                                else
                                {
                                    for (; i0 < c0s.Count; i0++)
                                    {
                                        Column c0 = c0s[i0];
                                        if (c.Header == c0.Header)
                                            break;
                                    }
                                    if (i0 >= c0s.Count)
                                        throw new Exception("Column[X=" + (i + 1) + "] '" + c.Header + "' has no match in the table.");
                                }
                            }
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
            }

            /// <summary>
            /// (!)Input columns must contain NULLs if any!
            /// NULL and empty-header columns are passed through.
            /// </summary>
            /// <param name="columns"></param>
            /// <param name="write"></param>
            /// <exception cref="Exception"></exception>
            void setColumns(IEnumerable<Column> columns, bool write)
            {
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

                Columns = new ReadOnlyCollection<Column>(columns.OrderBy(a => a.X).ToList());
                Columns.ForEach(a => { a.Table = this; });

                int x0 = 0;
                for (int i = 0; i < Columns.Count; i++)
                {
                    Column c = Columns[i];
                    if (c.X <= x0)
                        throw new Exception("Column[X=" + c.X + "] '" + c.Header + "' must have X>" + x0);
                    x0 = c.X;
                    for (int j = i + 1; j < Columns.Count; j++)
                    {
                        Column cj = Columns[j];
                        if (cj.X == c.X)
                            throw new Exception("Columns have the same X: '" + c.Header + "'[X=" + c.X + "] == '" + cj.Header + "'[X=" + cj.X + "]");
                        if (cj.Header == c.Header)
                            throw new Exception("Columns have the same headers: '" + c.Header + "'[X=" + c.X + "] == '" + cj.Header + "'[X=" + cj.X + "]");
                    }
                }

                if (Columns.Count < 1)
                    return;

                if (write)
                {
                    IRow r = Sheet._GetRow(1, true);
                    foreach (var column in Columns)
                    {
                        var c = r._GetCell(column.X, true);
                        c._SetValue(column.Header);
                        if (column.HeaderStyle != null)
                            c.CellStyle = column.HeaderStyle;
                    }
                }

                {//set data styles and types
                    var r2 = Sheet._GetRow(2, false);
                    bool r2created = false;
                    if (r2 == null)
                    {
                        r2 = Sheet._GetRow(2, true);
                        r2created = true;
                    }
                    Columns.Where(a => a.Style == null).ForEach(a =>
                    {
                        ICell c = r2._GetCell(a.X, false);
                        bool ccreated = false;
                        if (c == null)
                        {
                            c = r2._GetCell(a.X, true);
                            ccreated = true;
                        }
                        a.Style = c.CellStyle;
                        if (ccreated)
                            c._Remove(false);
                    });
                    Columns.Where(a => a.Type == null).ForEach(a =>
                    {
                        ICell c = r2._GetCell(a.X, false);
                        bool ccreated = false;
                        if (c == null)
                        {
                            c = r2._GetCell(a.X, true);
                            ccreated = true;
                        }
                        a.Type = c.CellType;
                        if (ccreated)
                            c._Remove(false);
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

            public void InsertColumn(int x, Column column, MoveRegionMode moveRegionMode)
            {
                Sheet._ShiftColumnsRight(x, 1, moveRegionMode);
                if (column == null)
                    return;
                column.X = x;
                Sheet._GetCell(1, x, true)._SetValue(column.Header);
                var cs = Columns.ToList();
                cs.Insert(column.X - 1, column);
                loadColumns();
                SetColumns(SetColumnMode.FindOrdered, cs);
            }

            public void RemoveColumn(Column column, MoveRegionMode moveRegionMode = null)
            {
                if (column.Table == null)
                    throw new Exception("Column is not initialized: Table is not set.");
                Sheet._ShiftColumnsLeft(column.X, 1, moveRegionMode);
                var cs = Columns.ToList();
                cs.RemoveAt(column.X - 1);
                loadColumns();
                SetColumns(SetColumnMode.FindOrdered, cs);
            }

            public class Column
            {
                public readonly string Header;
                public int X { get; internal set; } = -1;

                /// <summary>
                /// (!)Unregistered style will be registered when setting.
                /// </summary>
                public ICellStyle Style
                {
                    get
                    {
                        return style;
                    }
                    set
                    {
                        if (value == null)
                            return;
                        if (value.Index < 0)
                            value = Table.Excel.Workbook._GetRegisteredStyle(value);
                        style = value;
                        //Table?.Sheet.SetDefaultColumnStyle(X - 1, style);
                    }
                }
                ICellStyle style = null;
                public void ApplyStyle(ICellStyle style = null)
                {
                    if (style != null)
                        Style = style;
                    foreach (ICell c in GetDataCells(RowScope.WithCells))
                        c.CellStyle = Style;
                }

                public CellType? Type { get; set; } = null;
                public void ApplyType(CellType? type = null)
                {
                    if (type != null)
                        Type = type.Value;
                    if (Type != null)
                        foreach (ICell c in GetDataCells(RowScope.WithCells))
                            c.SetCellType(Type.Value);
                }

                readonly public ICellStyle HeaderStyle;

                public int GetWidth()
                {
                    return Table._.GetColumnWidth(X - 1);
                }

                public void SetWidth(int width)
                {
                    Table._._SetColumnWidth(X - 1, width);
                }

                public Table Table { get; internal set; } = null;

                /// <summary>
                /// (!)Until a created column is registered in Excel.Table.Columns, it is not initialized and cannot be used in most methods.
                /// </summary>
                /// <param name="header"></param>
                /// <param name="style"></param>
                /// <param name="type"></param>
                /// <param name="headerStyle"></param>
                public Column(string header, ICellStyle style = null, CellType? type = null, ICellStyle headerStyle = null/*???headerStyle is used only once so it seems to be pretty useless in constructor*/)
                {
                    if (string.IsNullOrWhiteSpace(header))
                        throw new Exception("Header cannot be empty or space.");
                    Header = header;
                    Style = style;
                    Type = type;
                    HeaderStyle = headerStyle;
                }

                /// <summary>
                /// (!)Until a created column is registered in Excel.Table.Columns, it is not initialized and cannot be used in most methods.
                /// </summary>
                /// <param name="headerStyle"></param>
                /// <param name="header"></param>
                /// <param name="style"></param>
                /// <param name="type"></param>
                /// <exception cref="Exception"></exception>
                public Column(ICellStyle headerStyle/*???headerStyle is used only once so it seems to be pretty useless in constructor*/, string header, ICellStyle style = null, CellType? type = null)
                {
                    if (string.IsNullOrWhiteSpace(header))
                        throw new Exception("Header cannot be empty or space.");
                    Header = header;
                    Style = style;
                    Type = type;
                    HeaderStyle = headerStyle;
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