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
        public void RemoveEmptyRows(bool removeIfEmptyCells)
        {
            Sheet._RemoveEmptyRows(removeIfEmptyCells);
        }

        public int GetLastRow(bool includeMerged)
        {
            return Sheet._GetLastRow(includeMerged);
        }

        public IEnumerable<IRow> GetRows(RowScope rowScope)
        {
            return Sheet._GetRows(rowScope);
        }

        public IRow GetLastRowWithCells()
        {
            return Sheet._GetLastRowWithCells();
        }
    }
}