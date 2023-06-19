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
            public class NamedValue
            {
                public Column Column { get; internal set; }
                public object Value { get; internal set; }
                public int X { get { return Column.X; } }
                public Func<ICell, bool> IsValueMatch = null;

                public NamedValue(Column column, object value, Func<ICell, bool> isValueMatch = null)
                {
                    if (column.Table == null)
                        throw new Exception("Column does not belong to a Table.");
                    Column = column;
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, string header, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(header);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, Func<string, bool> isHeaderMatch, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(isHeaderMatch);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }

                public NamedValue(Table table, Regex headerRegex, object value, Func<ICell, bool> isValueMatch = null)
                {
                    Column = table.GetColumn(headerRegex);
                    Value = value;
                    IsValueMatch = isValueMatch;
                }
            }
        }
    }
}