﻿//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Cliver
{
    public partial class Excel
    {
        public void RemoveEmptyRows(bool includeEmptyCellRows, bool shiftRowsBelow)
        {
            Sheet._RemoveEmptyRows(includeEmptyCellRows, shiftRowsBelow);
        }

        public enum RowScope
        {
            /// <summary>
            /// Returns only rows with at least one not empty cell.
            /// (!)Slow due to checking all the cells' values.
            /// </summary>
            NotEmpty,
            /// <summary>
            /// Returns only rows with cells.
            /// </summary>
            WithCells,
            /// <summary>
            /// Returns only rows existing as objects.
            /// </summary>
            NotNull,
            /// <summary>
            /// Returns all the rows withing the range with non-existing rows represented as NULL. 
            /// (!)Might return a huge pile of null and no-cell rows after the last not empty row.  
            /// </summary>
            IncludeNull,
            /// <summary>
            /// Returns all the rows withing the range with non-existing rows having been created.
            /// </summary>
            CreateIfNull
        }
        public IEnumerable<IRow> GetRows(RowScope rowScope)
        {
            return Sheet._GetRows(rowScope);
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

        public IRow GetRow(int y, bool createRow)
        {
            return Sheet._GetRow(y, createRow);
        }

        public IEnumerable<IRow> GetRowsInRange(RowScope rowScope, int y1 = 1, int? y2 = null)
        {
            return Sheet._GetRowsInRange(rowScope, y1, y2);
        }

        public IRow AppendRow<T>(IEnumerable<T> values)
        {
            return Sheet._AppendRow(values);
        }

        public IRow AppendRow(params string[] values)
        {
            return Sheet._AppendRow(values);
        }

        public IRow InsertRow<T>(int y, IEnumerable<T> values = null)
        {
            return Sheet._InsertRow(y, values);
        }

        public IRow InsertRow(int y, params string[] values)
        {
            return Sheet._InsertRow(y, values);
        }

        public IRow WriteRow<T>(int y, IEnumerable<T> values)
        {
            return Sheet._WriteRow(y, values);
        }

        public IRow WriteRow<T>(int y, params string[] values)
        {
            return Sheet._WriteRow(y, values);
        }

        public enum RemoveRowMode
        {
            ShiftRowsBelow = 1,
            ClearMerging = 2,
            /// <summary>
            /// (!)Done in a hacky way through Reflection so might change with POI update.
            /// (!)GetCell() might work incorrectly on such rows.
            /// </summary>
            PreserveCells = 4,
        }
        public IRow RemoveRow(int y, RemoveRowMode removeRowMode = 0)
        {
            return Sheet._RemoveRow(y, removeRowMode);
        }

        public void MoveRow(int y1, int y2, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
        {
            Sheet._MoveRow(y1, y2, onFormulaCellMoved, toSheet);
        }

        public void CopyRow(int y1, int y2, OnFormulaCellMoved onFormulaCellMoved = null, ISheet toSheet = null)
        {
            Sheet._CopyRow(y1, y2, onFormulaCellMoved, toSheet);
        }

        public void ShiftRowCellsRight(int y, int x1, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftRowCellsRight(y, x1, shift, onFormulaCellMoved);
        }

        public void ShiftRowCellsLeft(int y, int x1, int shift, OnFormulaCellMoved onFormulaCellMoved = null)
        {
            Sheet._ShiftRowCellsLeft(y, x1, shift, onFormulaCellMoved);
        }

        public void SetStyleInRow(ICellStyle style, bool createCells, int y)
        {
            Sheet._SetStyleInRow(style, createCells, y);
        }

        public void SetStyleInRowRange(ICellStyle style, bool createCells, int y1, int? y2 = null)
        {
            Sheet._SetStyleInRowRange(style, createCells, y1, y2);
        }

        public void ReplaceStyleInRowRange(ICellStyle style1, ICellStyle style2, int y1, int? y2 = null)
        {
            Sheet._ReplaceStyleInRowRange(style1, style2, y1, y2);
        }

        public void ClearStyleInRowRange(ICellStyle style, int y1, int? y2 = null)
        {
            Sheet._ClearStyleInRowRange(style, y1, y2);
        }

        public void AutosizeRowsInRange(int y1 = 1, int? y2 = null)
        {
            Sheet._AutosizeRowsInRange(y1, y2);
        }

        public void AutosizeRows()
        {
            Sheet._AutosizeRows();
        }

        public void ClearMergingInRow(int y)
        {
            Sheet._ClearMergingInRow(y);
        }

        public int GetLastColumnInRowRange(bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return Sheet._GetLastColumnInRowRange(includeMerged, y1, y2);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRow(bool includeMerged, int y)
        {
            return Sheet._GetLastNotEmptyColumnInRow(includeMerged, y);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastColumnInRow(bool includeMerged, int y)
        {
            return Sheet._GetLastColumnInRow(includeMerged, y);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="includeMerged"></param>
        /// <param name="y1"></param>
        /// <param name="y2"></param>
        /// <returns>1-based, otherwise 0</returns>
        public int GetLastNotEmptyColumnInRowRange(bool includeMerged, int y1 = 1, int? y2 = null)
        {
            return Sheet._GetLastNotEmptyColumnInRowRange(includeMerged, y1, y2);
        }

        public enum RowStyleMode
        {
            /// <summary>
            /// Set the row default style.
            /// </summary>
            Row = 1,
            /// <summary>
            /// Set style to the existing cells.
            /// </summary>
            ExistingCells = 2,
            /// <summary>
            /// Set style to all the cells with no gaps. When need, blank cells are created.
            /// </summary>
            NoGapCells = 4,
        }
        public void SetStyle(int y, ICellStyle style, RowStyleMode rowStyleMode)
        {
            Sheet._GetRow(y, true)._SetStyle(style, rowStyleMode);
        }
    }
}