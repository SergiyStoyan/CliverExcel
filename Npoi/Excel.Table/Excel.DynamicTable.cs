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
using System.Drawing;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Cliver
{
    public partial class Excel : IDisposable
    {
        public class DynamicTable : Excel.Table
        {
            public DynamicTable(ISheet sheet) : base(sheet)
            {
                HeaderStyle = Excel._._CreateUnregisteredStyle();
                var headerFont = Excel._._CloneUnregisteredFont(Sheet._GetCell(1, 1, true).CellStyle.GetFont(Excel._));
                headerFont.IsBold = true;
                headerFont = Excel._._GetRegisteredFont(headerFont);
                HeaderStyle.SetFont(headerFont);
                HeaderStyle = Excel._._GetRegisteredStyle(HeaderStyle);

                //BaseUri = new Uri(PathRoutines.GetFileDir(Excel.File) + "\\", UriKind.Absolute);
                //Excel._._SetHyperlinkBase(Excel.AbsoluteLinksHyperlinkBase);
            }

            public ICellStyle HeaderStyle;

            //public Uri BaseUri;

            /// <summary>
            /// optional
            /// </summary>
            /// <param name="headers"></param>
            public void SetHeaders(IEnumerable<string> headers)
            {
                List<Column> columns = headers.Select(a => new Column(a)).ToList();
                SetColumns(SetColumnMode.Override, columns);
                Sheet._SetStyleInRow(HeaderStyle, true, 1);
            }

            public void AppendRow(Dictionary<string, object> rowObject)
            {
                List<Cell> cells = new List<Cell>();
                foreach (var nv in rowObject)
                {
                    var c = GetColumn(nv.Key, false);
                    if (c == null)
                    {
                        c = new Column(HeaderStyle, nv.Key);
                        AddColumn(c, null);
                    }
                    cells.Add(new Cell(c, nv.Value));
                }
                AppendRow(cells);
            }
        }
    }
}
