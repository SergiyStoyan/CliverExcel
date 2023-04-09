//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using static Cliver.Excel;

namespace Cliver
{
    public partial class Style
    {
        public Style(ICellStyle style, Workbook workbook)
        {
            _ = style;
            Workbook = workbook;
        }
        public ICellStyle _ { get; private set; }
        public Workbook Workbook { get; private set; }
    }
}