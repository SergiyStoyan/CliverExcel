//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Extractor;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.Extractor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static Cliver.Excel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public IComment _SetComment(this ICell cell, string comment, string author = null, int paddingHeight = 2, int width = 3, IFont font = null, IFont authorFont = null)
        {
            cell.RemoveCellComment();//!!!adding multiple comments brings to error
            if (string.IsNullOrWhiteSpace(comment))
                return null;

            string @string = null;
            string author_ = null;
            if (!string.IsNullOrEmpty(author))
            {
                author_ = author + ":";
                @string = author_ + "\r\n";
            }
            @string += comment;
            var drawingPatriarch = /*cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch :*/ cell.Sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawingPatriarch.CreateAnchor(0, 0, 0, 0,
                    cell.ColumnIndex,
                    cell.RowIndex,
                    cell.ColumnIndex + width,
                    cell.RowIndex + Regex.Matches(@string, @"^", RegexOptions.Multiline).Count + paddingHeight
                    );
            IComment iComment = drawingPatriarch.CreateCellComment(anchor);
            List<RichTextStringFormattingRun> rtsfrs = font != null ? new List<RichTextStringFormattingRun> { new RichTextStringFormattingRun(0, @string.Length, font) } : null;
            if (!string.IsNullOrEmpty(author))
            {
                iComment.Author = author;
                if (authorFont == null)
                {
                    iComment.String = cell.Sheet.Workbook._GetRichTextString(@string, rtsfrs);
                    var f = _GetRichTextStringFormattingRuns(cell.Sheet.Workbook, iComment.String).FirstOrDefault(a => a.Font != null)?.Font;
                    if (f == null)//(!)on XSSFWorkbook, RichTextString can have no FormattingRuns or have FormattingRuns with Font=NULL
                        f = cell.Sheet.Workbook._GetCommentDefaultFont();
                    authorFont = cell.Sheet.Workbook._CloneUnregisteredFont(f);
                    authorFont.IsBold = true;
                    authorFont = cell.Sheet.Workbook._GetRegisteredFont(authorFont);
                }
                iComment.String.ApplyFont(0, author_.Length, authorFont);
            }
            else
                iComment.String = cell.Sheet.Workbook._GetRichTextString(comment, rtsfrs);
            cell.CellComment = iComment;

            return cell.CellComment;
        }

        static public IComment _AppendOrSetComment(this ICell cell, string comment, string author = null, int paddingHeight = 0, int width = 3, string delimiter = "\r\n", IFont font = null, IFont authorFont = null)
        {
            if (string.IsNullOrWhiteSpace(comment))
                return cell?.CellComment;

            string string1 = cell?.CellComment?.String?.String;
            if (string.IsNullOrEmpty(string1))
                return cell._SetComment(comment, author, paddingHeight < 2 ? 2 : paddingHeight, width, font);//(!)the first comment goes with altered paddingHeight

            List<RichTextStringFormattingRun> rtsfrs = cell.Sheet.Workbook._GetRichTextStringFormattingRuns(cell.CellComment.String).ToList();
            string string2 = delimiter;
            if (!string.IsNullOrEmpty(author))
            {
                if (authorFont == null)
                {
                    if (font == null)
                    {
                        authorFont = rtsfrs.Select(a => a.Font).FirstOrDefault(a => a?.IsBold == true);
                        if (authorFont == null)
                        {
                            IFont f = rtsfrs.Select(a => a.Font).FirstOrDefault(a => a?.IsBold == false);
                            if (f == null)//(!)on XSSFWorkbook, RichTextString can have no FormattingRuns or have FormattingRuns with Font=NULL
                                f = cell.Sheet.Workbook._GetCommentDefaultFont();
                            authorFont = cell.Sheet.Workbook._CloneUnregisteredFont(f);
                            authorFont.IsBold = true;
                            authorFont = cell.Sheet.Workbook._GetRegisteredFont(authorFont);
                        }
                    }
                    else
                    {
                        authorFont = cell.Sheet.Workbook._CloneUnregisteredFont(font);
                        authorFont.IsBold = true;
                        authorFont = cell.Sheet.Workbook._GetRegisteredFont(authorFont);
                    }
                }
                string author_ = author + ":";
                rtsfrs.Add(new RichTextStringFormattingRun(string1.Length + delimiter.Length, string1.Length + delimiter.Length + author_.Length, authorFont));
                string2 += author_ + "\r\n";
            }
            string2 += comment;
            string @string = string1 + string2;
            if (font != null)
                rtsfrs.Add(new RichTextStringFormattingRun(@string.Length - comment.Length, @string.Length, font));
            var drawingPatriarch = /*cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch :*/ cell.Sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawingPatriarch.CreateAnchor(
                cell.CellComment.ClientAnchor.Dx1,
                cell.CellComment.ClientAnchor.Dy1,
                cell.CellComment.ClientAnchor.Dx2,
                cell.CellComment.ClientAnchor.Dy2,
                cell.CellComment.ClientAnchor.Col1,
                cell.CellComment.ClientAnchor.Row1,
                cell.CellComment.ClientAnchor.Col2,
                cell.CellComment.ClientAnchor.Row2 + Regex.Matches(string2, @"^", RegexOptions.Multiline).Count + paddingHeight
            );
            cell.RemoveCellComment();
            IComment iComment = drawingPatriarch.CreateCellComment(anchor);
            if (author != null)//(!)set the last author
                iComment.Author = author;
            iComment.String = cell.Sheet.Workbook._GetRichTextString(@string, rtsfrs);
            cell.CellComment = iComment;

            return cell.CellComment;
        }
        //static public IComment _SetComment(this ICell cell, string comment, string author = null, int paddingHeight = 3, int width = 3)
        //{
        //    cell.RemoveCellComment();//!!!adding multiple comments brings to error            
        //    return _AppendOrSetComment(cell, comment, author, "\r\n\r\n", paddingHeight, width);
        //}
        //static public IComment _AppendOrSetComment(this ICell cell, string comment, string author = null, string delimiter = "\r\n\r\n", int paddingHeight = 3, int width = 3)
        //{
        //    if (string.IsNullOrWhiteSpace(comment))
        //        return cell?.CellComment;

        //    string string1 = cell?.CellComment?.String?.String;
        //    var drawingPatriarch = cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch : cell.Sheet.CreateDrawingPatriarch();
        //    IClientAnchor anchor;
        //    int string2Linescount = Regex.Matches(comment, @"^", RegexOptions.Multiline).Count + paddingHeight;
        //    if (string.IsNullOrEmpty(string1))
        //        anchor = drawingPatriarch.CreateAnchor(0, 0, 0, 0,
        //            cell.ColumnIndex,
        //            cell.RowIndex,
        //            cell.ColumnIndex + width,
        //            cell.RowIndex + string2Linescount
        //            );
        //    else
        //        anchor = drawingPatriarch.CreateAnchor(0, 0, 0, 0,
        //             cell.CellComment.ClientAnchor.Col1,
        //             cell.CellComment.ClientAnchor.Row1,
        //             cell.CellComment.ClientAnchor.Col2,
        //             cell.CellComment.ClientAnchor.Row2 + string2Linescount
        //             );
        //    IComment iComment = drawingPatriarch.CreateCellComment(anchor);
        //    if (author != null)//(!)set the last author
        //        iComment.Author = author;
        //    if (!string.IsNullOrEmpty(string1))
        //        string1 += delimiter;
        //    string string2 = null;
        //    string author_ = null;
        //    if (!string.IsNullOrEmpty(author))
        //    {
        //        author_ = author + ":";
        //        string2 += author_ + "\r\n";
        //    }
        //    string2 += comment;
        //    List<RichTextStringFormattingRun> rtsfrs;
        //    if (!string.IsNullOrEmpty(string1))
        //    {
        //        rtsfrs = cell.Sheet.Workbook._GetRichTextStringFormattingRuns(cell.CellComment.String).ToList();
        //        iComment.String = cell.Sheet.Workbook._GetRichTextString(string1 + string2, rtsfrs);
        //    }
        //    else
        //    {
        //        iComment.String = cell.Sheet.Workbook._GetRichTextString(string2, null);
        //        rtsfrs = cell.Sheet.Workbook._GetRichTextStringFormattingRuns(iComment.String).ToList();
        //    }
        //    if (!string.IsNullOrEmpty(author))
        //    {
        //        IFont fb = rtsfrs.Select(a => a.Font).FirstOrDefault(a => a.IsBold);
        //        if (fb == null)
        //        {
        //            IFont f = rtsfrs.Select(a => a.Font).FirstOrDefault(a => !a.IsBold);
        //            fb = cell.Sheet.Workbook._CloneUnregisteredFont(f);
        //            fb.IsBold = true;
        //            fb = cell.Sheet.Workbook._GetRegisteredFont(fb);
        //        }
        //        iComment.String.ApplyFont(string1.Length, string1.Length + author_.Length, fb);
        //    }
        //    cell.RemoveCellComment();
        //    cell.CellComment = iComment;

        //    return cell.CellComment;
        //}

        //public static IComment _CopyComment(this ICell cell, int y2, int x2)
        //{
        //    IComment comment = cell.CellComment;
        //    if (comment == null)
        //        return null;
        //    var drawingPatriarch = /*cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch :*/ cell.Sheet.CreateDrawingPatriarch();
        //    (int Y, int X) shift = (y2 - cell._Y(), x2 - cell._X());
        //    IClientAnchor anchor2 = drawingPatriarch.CreateAnchor(0, 0, 0, 0
        //        , comment.ClientAnchor.Col1 + shift.X
        //        , comment.ClientAnchor.Row1 + shift.Y
        //        , comment.ClientAnchor.Col2 + shift.X
        //        , comment.ClientAnchor.Row2 + shift.Y
        //        );
        //    //IClientAnchor anchor2 = drawingPatriarch.CreateAnchor(0, 0, 0, 0
        //    //    , x2
        //    //    , y2
        //    //    , x2 + comment.ClientAnchor.Col2 - comment.ClientAnchor.Col1
        //    //    , y2 + comment.ClientAnchor.Row2 - comment.ClientAnchor.Row1
        //    //    );
        //    IComment comment2 = drawingPatriarch.CreateCellComment(anchor2);
        //    if (comment.Author != null)
        //        comment2.Author = comment.Author;
        //    comment2.String = comment.String.Copy();
        //    return comment2;
        //}
    }
}