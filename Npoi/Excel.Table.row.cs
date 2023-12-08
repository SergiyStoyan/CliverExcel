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
            /// Intended for easier filtering and CRUD operations. So it represents values only.
            /// </summary>
            public class Row
            {
                public virtual Table Table { get; private set; }
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

                public Row(Table table)
                {
                    Table = table;
                }

                public Row(Table table, params Cell[] cells) : this(table, (IEnumerable<Cell>)cells)
                { }

                public Row(Table table, IEnumerable<Cell> cells)
                {
                    Table = table;
                    foreach (Cell c in cells)
                        Cells[c.X] = c;
                }

                public Row(Table table, IRow iRow)
                {
                    Table = table;
                    foreach (Column c in table.Columns)
                        Cells[c.X] = new Cell(c, iRow._GetCell(c.X, false)?._GetValue());
                }
            }

            public IEnumerable<Row> GetRows(IEnumerable<IRow> iRows)
            {
                return iRows.Select(a => new Row(this, a));
            }

            public IRow AppendRow(Row row)
            {
                return AppendRow(row.Cells.Values);
            }

            public IRow InsertRow(int y, Row row)
            {
                return InsertRow(y, row.Cells.Values);
            }

            public IRow WriteRow(int y, Row row)
            {
                return WriteRow(y, row.Cells.Values);
            }
        }
    }
}