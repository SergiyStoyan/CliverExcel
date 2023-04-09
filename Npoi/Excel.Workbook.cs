//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Cliver
{
    public partial class Workbook
    {
        public IEnumerable<Sheet> GetSheets()
        {
            for (int i = 0; i < _.ActiveSheetIndex; i++)
            {
                yield return new Sheet(_.GetSheetAt(i));
            }
        }
    }
}