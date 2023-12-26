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
using System.Collections.Specialized;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public partial class Table
        {
            /// <summary>
            /// Intended for easier filtering and CRUD operations.
            /// </summary>
            public class Row<T> where T : Table
            {
                public virtual T Table { get; private set; }
                public SortedDictionary<int, Cell> Cells { get; private set; } = new SortedDictionary<int, Cell>();

                public Cell this[Column column]
                {
                    get
                    {
                        Cells.TryGetValue(column.X, out Cell cell);
                        return cell;
                    }
                    set
                    {
                        Cells[column.X] = value;
                    }
                }

                public Row(T table)
                {
                    Table = table;
                }

                public Row(T table, params Cell[] cells) : this(table, (IEnumerable<Cell>)cells)
                { }

                public Row(T table, IEnumerable<Cell> cells)
                {
                    Table = table;
                    foreach (Cell c in cells)
                        Cells[c.X] = c;
                }

                public Row(T table, IRow iRow, Get get)
                {
                    Table = table;
                    foreach (Column column in table.Columns)
                    {
                        ICell c = iRow._GetCell(column.X, false);
                        Cells[column.X] = new Cell(column,
                            get.HasFlag(Get.Value) ? c?._GetValue() : null,
                            get.HasFlag(Get.Value) ? c?.CellStyle : null,
                            get.HasFlag(Get.Type) ? c?.CellType : null,
                            get.HasFlag(Get.Link) ? c?._GetLink() : null
                            );
                    }
                }
            }

            public enum Get
            {
                Value,
                Style,
                Type,
                Link,
            }

            public IEnumerable<Row<Table>> GetRows(IEnumerable<IRow> iRows, Get get)
            {
                return iRows.Select(a => new Row<Table>(this, a, get));
            }

            public IRow AppendRow<T>(Row<T> row) where T : Table
            {
                return AppendRow(row.Cells.Values);
            }

            public IRow InsertRow<T>(int y, Row<T> row) where T : Table
            {
                return InsertRow(y, row.Cells.Values);
            }

            public IRow WriteRow<T>(int y, Row<T> row) where T : Table
            {
                return WriteRow(y, row.Cells.Values);
            }
        }
    }
}