//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;

namespace Cliver
{
    public partial class Excel
    {
        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            Sheet._ReplaceStyle(style1, style2);
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            Sheet._SetStyle(style, createCells);
        }

        public void UnsetStyle(ICellStyle style)
        {
            Sheet._UnsetStyle(style);
        }

        public Range NewRange(int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return Sheet._NewRange(y1, x1, y2, x2);
        }
    }
}