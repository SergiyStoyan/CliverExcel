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
        public void RemoveEmptyRows(bool includeEmptyCellRows, bool shiftRemainingRows)
        {
            Sheet._RemoveEmptyRows(includeEmptyCellRows, shiftRemainingRows);
        }

        public enum LastRowCondition
        {
            /// <summary>
            /// (!)Considerably slow due to checking all the cells' values
            /// </summary>
            NotEmpty,
            /// <summary>
            /// Row with cells.
            /// </summary>
            HasCells,
            /// <summary>
            /// Row existing as an object.
            /// </summary>
            NotNull,
        }

        public int GetLastRow(LastRowCondition lastRowCondition, bool includeMerged)
        {
            return Sheet._GetLastRow(lastRowCondition, includeMerged);
        }

        public IEnumerable<IRow> GetRows(RowScope rowScope)
        {
            return Sheet._GetRows(rowScope);
        }
    }
}