//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Linq;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.OpenXml4Net.OPC;
using NPOI;

namespace Cliver
{
    public partial class Excel
    {
        /// <summary>
        /// (!) 1-based
        /// </summary>
        public class Range
        {
            public int X1 = 1;
            public int? X2 = null;
            public int Y1 = 1;
            public int? Y2 = null;

            internal Range(ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
            {
                Sheet = sheet;
                Y1 = y1;
                Y2 = y2;
                X1 = x1;
                X2 = x2;
            }

            public ISheet Sheet;

            public ICell GetFirstCell(bool createCell)
            {
                return Sheet._GetCell(Y1, X1, createCell);
            }

            public string GetStringAddress()
            {
                return CellReference.ConvertNumToColString(X1 - 1) + Y1 + ":" + CellReference.ConvertNumToColString(X2 != null ? X2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.LastColumnIndex) + Y2;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>(!) 0-based</returns>
            public CellRangeAddress GetCellRangeAddress()
            {
                return new CellRangeAddress(Y1 - 1, Y2 != null ? Y2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.MaxRows - 1, X1 - 1, X2 != null ? X2.Value - 1 : Sheet.Workbook.SpreadsheetVersion.LastColumnIndex);
            }

            public void Clear(bool clearMerging, bool removeComment = true)
            {
                if (clearMerging)
                    ClearMerging();

                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, false);
                    if (row == null)
                        continue;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                        row._GetCell(x, false)?._Remove(removeComment);
                }
            }

            public void ClearMerging()
            {
                CellRangeAddress cra = GetCellRangeAddress();
                for (int i = Sheet.MergedRegions.Count - 1; i >= 0; i--)
                    if (Sheet.MergedRegions[i].Intersects(cra))
                        Sheet.RemoveMergedRegion(i);
            }

            public void Merge(bool clearOldMerging = false)
            {
                if (clearOldMerging)
                    ClearMerging();
                Sheet.AddMergedRegion(GetCellRangeAddress());
            }

            public bool Contains(CellAddress cellAddress)
            {
                return Contains(cellAddress.Row + 1, cellAddress.Column + 1);
            }

            public bool Contains(ICell c)
            {
                return Contains(c.RowIndex + 1, c.ColumnIndex + 1);
            }

            public bool Contains(int y, int x)
            {
                return y >= Y1 && (Y2 == null || y <= Y2.Value) && x >= X1 && (X2 == null || x <= X2.Value);
            }

            public bool Contains(Range r)
            {
                return r.Y1 >= Y1 && (Y2 == null || r.Y1 <= Y2.Value) && r.X1 >= X1 && (X2 == null || r.X2 <= X2.Value);
            }

            public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, false);
                    if (row == null)
                        continue;
                    if (Y1 == 1 && Y2 == null
                        && row.RowStyle?.Index == style1.Index
                        )
                        row.RowStyle = style2;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell c = row._GetCell(x, false);
                        if (c != null && c.CellStyle?.Index == style1.Index)
                            c.CellStyle = style2;
                    }
                }
            }

            public void SetStyle(ICellStyle style, bool createCells)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet._GetRow(y, createCells);
                    if (row == null)
                        continue;
                    if (Y1 == 1 && Y2 == null)
                        row.RowStyle = style;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell c = row._GetCell(x, createCells);
                        if (c != null)
                            c.CellStyle = style;
                    }
                }
            }

            public void UnsetStyle(ICellStyle style)
            {
                ReplaceStyle(style, null);
            }

            public void SetAlteredStyles<T>(T alterationKey, Excel.StyleCache.AlterStyle<T> alterStyle, CellScope cellScope/*, bool reuseUnusedStyle = false*/) where T : Excel.StyleCache.IKey
            {
                foreach (var c in GetCells(cellScope))
                    c?._SetAlteredStyle(alterationKey, alterStyle/*, reuseUnusedStyle*/);
            }

            public IEnumerable<ICell> GetCells(CellScope cellScope)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                int maxX;
                switch (cellScope)
                {
                    case CellScope.NotEmpty:
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            if (r == null)
                                continue;
                            maxX = X2 < r.LastCellNum ? X2.Value : r.LastCellNum;
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, false);
                                if (!string.IsNullOrWhiteSpace(c._GetValueAsString()))
                                    yield return c;
                            }
                        }
                        break;
                    case CellScope.NotNull:
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            if (r == null)
                                continue;
                            maxX = X2 < r.LastCellNum ? X2.Value : r.LastCellNum;
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, false);
                                if (c != null)
                                    yield return c;
                            }
                        }
                        break;
                    case CellScope.IncludeNull:
                        if (X2 != null)
                            maxX = X2.Value;
                        else
                        {
                            maxX = 0;
                            for (int y = Y1; y <= maxY; y++)
                            {
                                var r = Sheet._GetRow(y, false);
                                if (r != null && maxX < r.LastCellNum)
                                    maxX = r.LastCellNum;
                            }
                        }
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, false);
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r?._GetCell(x, false);
                                yield return c;
                            }
                        }
                        break;
                    case CellScope.CreateIfNull:
                        if (X2 != null)
                            maxX = X2.Value;
                        else
                        {
                            maxX = 0;
                            for (int y = Y1; y <= maxY; y++)
                            {
                                var r = Sheet._GetRow(y, false);
                                if (r != null && maxX < r.LastCellNum)
                                    maxX = r.LastCellNum;
                            }
                        }
                        for (int y = Y1; y <= maxY; y++)
                        {
                            var r = Sheet._GetRow(y, true);
                            for (int x = X1; x <= maxX; x++)
                            {
                                var c = r._GetCell(x, true);
                                yield return c;
                            }
                        }
                        break;
                    default:
                        throw new Exception("Unknown option: " + cellScope.ToString());
                }
            }

            ICell[][] copyCutRange(bool cut)
            {
                int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                ICell[][] rangeCells = new ICell[maxY - Y1 + 1][];
                for (int y = Y1; y <= maxY; y++)
                {
                    IRow row = Sheet.GetRow(y - 1);
                    if (row == null)
                        continue;
                    int maxX = X2 != null ? X2.Value : row.LastCellNum;
                    ICell[] rowCells = new ICell[maxX];
                    for (int x = X1; x <= maxX; x++)
                    {
                        ICell cell = row.GetCell(x - 1);
                        rowCells[x - X1] = cell;
                        if (cut)
                            row.RemoveCell(cell);
                    }
                    if (cut && X1 == 1 && X2 == null)
                        Sheet.RemoveRow(row);
                    rangeCells[y - Y1] = rowCells;
                }
                return rangeCells;
            }

            public ICell[][] Copy()
            {
                return copyCutRange(false);
            }

            public ICell[][] Cut()
            {
                return copyCutRange(true);
            }

            public void Move(int toY, int toX, Excel.CopyCellMode copyCellMode = null)
            {
                PasteRange(Cut(), toY, toX, copyCellMode);
            }

            public void Copy(int toY, int toX, Excel.CopyCellMode copyCellMode = null)
            {
                PasteRange(Copy(), toY, toX, copyCellMode);
            }

            public void SetComment(string comment, string author = null)
            {
                throw new System.Exception("TBD");
                //if (comment == null)
                //{
                //    int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                //    for (int y = Y1; y <= maxY; y++)
                //    {
                //        IRow row = Sheet._GetRow(y, false);
                //        if (row == null)
                //            continue;
                //        int maxX = X2 != null ? X2.Value : row.LastCellNum;
                //        for (int x = X1; x <= maxX; x++)
                //            row._GetCell(x, false)?.RemoveCellComment();
                //    }
                //}
                //else
                //{
                //    var creationHelper = Sheet.Workbook.GetCreationHelper();
                //    var richTextString = creationHelper.CreateRichTextString(comment);
                //    var clientAnchor = creationHelper.CreateClientAnchor();
                //    //clientAnchor.Col1 = cell.ColumnIndex + 1;
                //    //clientAnchor.Col2 = cell.ColumnIndex + 3;
                //    //clientAnchor.Row1 = cell.RowIndex + 1;
                //    //clientAnchor.Row2 = cell.RowIndex + 5;
                //    var drawingPatriarch = Sheet.CreateDrawingPatriarch();

                //    int maxY = Y2 != null ? Y2.Value : Sheet.LastRowNum + 1;
                //    for (int y = Y1; y <= maxY; y++)
                //    {
                //        IRow row = Sheet._GetRow(y, false);
                //        if (row == null)
                //            continue;
                //        int maxX = X2 != null ? X2.Value : row.LastCellNum;
                //        for (int x = X1; x <= maxX; x++)
                //        {
                //            ICell cell = row._GetCell(x, true);
                //            IComment iComment = drawingPatriarch.CreateCellComment(clientAnchor);
                //            iComment.String = richTextString;
                //            if (!string.IsNullOrWhiteSpace(author))
                //                iComment.Author = author;
                //            cell.CellComment = iComment;
                //        }
                //    }
                //}
            }

            public void RemoveComments()
            {
                //Sheet.GetCellComments().Where(a => IsIn(a.Key)).ToList().ForEach(a => Sheet._GetCell(a.Key, false).RemoveCellComment());
                GetCells(CellScope.NotNull).ForEach(a => a.RemoveCellComment());
            }

            public void RemoveImages(ImageLocationType imageLocationType)
            {
                var ps = GetPictures(imageLocationType).ToList();
                if (ps.Count < 1)
                    return;
                var drawing = Sheet.CreateDrawingPatriarch();
                if (drawing is XSSFDrawing xssfDrawing)
                {
                    Dictionary<POIXMLDocumentPart, HashSet<string>> parts2embedIds = new Dictionary<POIXMLDocumentPart, HashSet<string>>();
                    foreach (var sh in Sheet.Workbook._GetSheets())
                    {
                        var ctD = ((XSSFDrawing)sh.CreateDrawingPatriarch())?.GetCTDrawing();
                        for (int ai = ctD.CellAnchors.Count - 1; ai >= 0; ai--)
                        {
                            var embedId = ctD.CellAnchors[ai]?.picture?.blipFill?.blip?.embed;
                            if (string.IsNullOrEmpty(embedId))
                                continue;
                            var pn = xssfDrawing.GetRelationById(embedId);
                            if (pn == null)
                                continue;
                            if (!parts2embedIds.TryGetValue(pn, out var embedIds))
                            {
                                embedIds = new HashSet<string>();
                                parts2embedIds[pn] = embedIds;
                            }
                            embedIds.Add(embedId);
                        }
                    }

                    var ctDrawing = xssfDrawing.GetCTDrawing();
                    foreach (XSSFPicture p in ps)
                    {
                        var ctP = p.GetCTPicture();
                        var embedId = ctP?.blipFill?.blip?.embed;
                        if (string.IsNullOrEmpty(embedId))
                            continue;

                        for (int i = ctDrawing.CellAnchors.Count - 1; i >= 0; i--)
                            if (ctDrawing.CellAnchors[i].picture?.blipFill?.blip?.embed?.Equals(embedId) == true)
                                ctDrawing.CellAnchors.RemoveAt(i);
                        var pp = xssfDrawing.GetPackagePart();
                        pp.RemoveRelationship(embedId);

                        var pn = xssfDrawing.GetRelationById(embedId);
                        if (pn != null)
                        {
                            var embedIds = parts2embedIds[pn];
                            embedIds.Remove(embedId);
                            if (embedIds.Count <= 0)//delete the image itself only when it has no reference remaining
                                pp.Package.DeletePartRecursive(pn.GetPackagePart().PartName);
                        }
                    }
                }
                else if (drawing is HSSFPatriarch hssfDrawing)
                {
                    foreach (HSSFPicture p in ps)
                        hssfDrawing.RemoveShape(p);
                }
                else
                    throw new Exception("Unsupported type: " + drawing.GetType());
            }
            public void RemoveImages3(ImageLocationType imageLocationType)
            {
                var ps = GetPictures(imageLocationType).ToList();
                if (ps.Count < 1)
                    return;
                var drawing = Sheet.CreateDrawingPatriarch();
                if (drawing is XSSFDrawing xssfDrawing)
                {
                    var ctDrawing = xssfDrawing.GetCTDrawing();
                    var xssfWorkbook = Sheet.Workbook as XSSFWorkbook;
                    foreach (XSSFPicture p in ps)
                    {
                        var ctP = p.GetCTPicture();
                        var embedId = ctP?.blipFill?.blip?.embed;
                        if (string.IsNullOrEmpty(embedId))
                            continue;

                        // remove only the anchor node that corresponds to this picture instance
                        if (ctDrawing != null)
                        {
                            for (int i = ctDrawing.CellAnchors.Count - 1; i >= 0; i--)
                            {
                                var anchor = ctDrawing.CellAnchors[i];
                                // remove the anchor only when it contains the very same CT picture object
                                if (ReferenceEquals(anchor.picture, ctP))
                                    ctDrawing.CellAnchors.RemoveAt(i);
                            }
                        }

                        // check whether the embedId is still referenced anywhere in the workbook
                        bool stillUsed = false;
                        if (xssfWorkbook != null)
                        {
                            for (int si = 0; si < xssfWorkbook.NumberOfSheets && !stillUsed; si++)
                            {
                                var sh = xssfWorkbook.GetSheetAt(si) as XSSFSheet;
                                if (sh == null)
                                    continue;
                                var dr = (XSSFDrawing)sh.CreateDrawingPatriarch();
                                var ctDr = dr?.GetCTDrawing();
                                if (ctDr == null)
                                    continue;
                                for (int ai = ctDr.CellAnchors.Count - 1; ai >= 0; ai--)
                                {
                                    var a = ctDr.CellAnchors[ai];
                                    if (a?.picture?.blipFill?.blip?.embed == embedId)
                                    {
                                        stillUsed = true;
                                        break;
                                    }
                                }
                            }
                        }

                        // if no remaining references, remove the relationship and delete the image part
                        if (!stillUsed)
                        {
                            var pp = xssfDrawing.GetPackagePart();
                            pp.RemoveRelationship(embedId);
                            var pn = xssfDrawing.GetRelationById(embedId)?.GetPackagePart()?.PartName;
                            if (pn != null)
                                pp.Package.DeletePartRecursive(pn);
                        }
                    }
                }
                else if (drawing is HSSFPatriarch hssfDrawing)
                {
                    foreach (HSSFPicture p in ps)
                        hssfDrawing.RemoveShape(p);
                }
                else
                    throw new Exception("Unsupported type: " + drawing.GetType());
            }

            public IEnumerable<Image> GetImages(ImageLocationType imageLocationType)
            {
                foreach (IPicture p in GetPictures(imageLocationType))
                {
                    var a = p.ClientAnchor;
                    IPictureData pictureData = p.PictureData;
                    yield return new Image { Data = pictureData.Data, Name = null, Type = pictureData.PictureType, X = a.Col1, Y = a.Row1/*, Anchor = a*/ };
                }
            }

            public IEnumerable<IPicture> GetPictures(ImageLocationType imageLocationType)
            {
                IEnumerable<IPicture> pictures;
                if (Sheet.Workbook is XSSFWorkbook xSSFWorkbook)
                {
                    XSSFDrawing dp = (XSSFDrawing)Sheet.CreateDrawingPatriarch();
                    pictures = dp.GetShapes().Where(a => a is IPicture).Select(a => (IPicture)a);
                }
                else if (Sheet.Workbook is HSSFWorkbook hWorkbook)
                {
                    HSSFPatriarch dp = (HSSFPatriarch)Sheet.CreateDrawingPatriarch();
                    pictures = dp.GetShapes().Where(a => a is IPicture).Select(a => (IPicture)a);
                }
                else
                    throw new Exception("Unsupported workbook type: " + Sheet.Workbook.GetType().FullName);

                switch (imageLocationType)
                {
                    case ImageLocationType.AnchorTopLeft:
                        foreach (IPicture p in pictures)
                        {
                            var a = p.ClientAnchor;
                            if (Contains(a.Row1 + 1, a.Col1 + 1))
                                yield return p;
                        }
                        break;
                    case ImageLocationType.WithinAnchor:
                        foreach (IPicture p in pictures)
                        {
                            var a = p.ClientAnchor;
                            if (Sheet._GetRange(a.Row1 + 1, a.Col1 + 1, a.Row2 + 1, a.Col2 + 1).Contains(this))
                                yield return p;
                        }
                        break;
                    case ImageLocationType.WithinRange:
                        foreach (IPicture p in pictures)
                        {
                            var a = p.ClientAnchor;
                            if (Contains(Sheet._GetRange(a.Row1 + 1, a.Col1 + 1, a.Row2 + 1, a.Col2 + 1)))
                                yield return p;
                        }
                        break;
                    default:
                        throw new Exception("Unknown " + nameof(imageLocationType) + ": " + imageLocationType);
                }
            }

            /// <summary>
            /// !!!it is a bug in NPOI-2.7 that Resize() changes picture's anchor ignoring AnchorType. 
            /// So, the pictures that belong to the cell should be rather filtered by the top-left anchor.
            /// </summary>
            public enum ImageLocationType
            {
                /// <summary>
                /// the picture's anchor Top-Left is within the range
                /// </summary>
                AnchorTopLeft,
                /// <summary>
                /// the range is within the picture's anchor (the anchors covers the range)
                /// </summary>
                WithinAnchor,
                /// <summary>
                /// picture's anchor is within the range (Top-Left and Bottom-Right are within the range)
                /// </summary>
                WithinRange,
            }
        }
    }
}
