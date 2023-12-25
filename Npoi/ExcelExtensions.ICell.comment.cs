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
        static public IComment _SetComment(this ICell cell, string comment, Excel.CommentStyle commentStyle = null)
        {
            cell.RemoveCellComment();//!!!adding multiple comments brings to error
            if (string.IsNullOrWhiteSpace(comment))
                return null;

            Excel.CommentStyle s = commentStyle != null ? commentStyle : cell.Sheet.Workbook._Excel().DefaultCommentStyle;

            string @string = null;
            string author_ = string.Empty;
            if (!string.IsNullOrEmpty(s.Author))
            {
                author_ = s.Author + ":";
                @string = author_ + s.AuthorDelimiter;
            }
            @string += comment;
            var drawingPatriarch = /*cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch :*/ cell.Sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawingPatriarch.CreateAnchor(
                0,
                cell.RowIndex == 0 ? 40 : 0/*to avoid bad representation*/,
                0,
                0,
                cell.ColumnIndex,
                cell.RowIndex,
                cell.ColumnIndex + s.Columns,
                cell.RowIndex + Regex.Matches(@string, @"^", RegexOptions.Multiline).Count + s.PaddingRows
                );
            IComment iComment = drawingPatriarch.CreateCellComment(anchor);//!!!due to NPOI implementation, it sets the comment to the cell
            List<RichTextStringFormattingRun> rtsfrs = new List<RichTextStringFormattingRun>();
            if (!string.IsNullOrEmpty(s.Author))
            {
                iComment.Author = s.Author;
                rtsfrs.Add(new RichTextStringFormattingRun(0, author_.Length, s.AuthorFont));
            }
            rtsfrs.Add(new RichTextStringFormattingRun(author_.Length, @string.Length, s.Font));
            iComment.String = cell.Sheet.Workbook._GetRichTextString(@string, rtsfrs);
            cell.CellComment = iComment;//!!!due to NPOI implementation, it is already set

            return cell.CellComment;
        }

        static public IComment _AppendOrSetComment(this ICell cell, string comment, Excel.CommentStyle commentStyle = null)
        {
            if (string.IsNullOrWhiteSpace(comment))
                return cell?.CellComment;

            Excel.CommentStyle s = commentStyle != null ? commentStyle : cell.Sheet.Workbook._Excel().DefaultCommentStyle;

            string string1 = cell?.CellComment?.String?.String;
            if (string.IsNullOrEmpty(string1))
                return cell._SetComment(comment, s);

            List<RichTextStringFormattingRun> rtsfrs = cell.Sheet.Workbook._GetRichTextStringFormattingRuns(cell.CellComment.String).ToList();
            string string2 = s.AppendDelimiter;
            if (!string.IsNullOrEmpty(s.Author))
            {
                string author_ = s.Author + ":";
                rtsfrs.Add(new RichTextStringFormattingRun(string1.Length + s.AppendDelimiter.Length, string1.Length + s.AppendDelimiter.Length + author_.Length, s.AuthorFont));
                string2 += author_ + s.AuthorDelimiter;
            }
            string2 += comment;
            rtsfrs.Add(new RichTextStringFormattingRun(string1.Length + string2.Length - comment.Length, string1.Length + string2.Length, s.Font));
            var drawingPatriarch = /*cell.Sheet.DrawingPatriarch != null ? cell.Sheet.DrawingPatriarch :*/ cell.Sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawingPatriarch.CreateAnchor(
                cell.CellComment.ClientAnchor.Dx1,
                cell.CellComment.ClientAnchor.Dy1,
                cell.CellComment.ClientAnchor.Dx2,
                cell.CellComment.ClientAnchor.Dy2,
                cell.CellComment.ClientAnchor.Col1,
                cell.CellComment.ClientAnchor.Row1,
                cell.CellComment.ClientAnchor.Col2,
                cell.CellComment.ClientAnchor.Row2 + Regex.Matches(string2, @"^", RegexOptions.Multiline).Count + s.AppendPaddingRows
            );
            cell.RemoveCellComment();
            IComment iComment = drawingPatriarch.CreateCellComment(anchor);//!!!due to NPOI implementation, it sets the comment to the cell
            if (s.Author != null)//(!)set the last author
                iComment.Author = s.Author;
            iComment.String = cell.Sheet.Workbook._GetRichTextString(string1 + string2, rtsfrs);
            cell.CellComment = iComment;//!!!due to NPOI implementation, it is already set;

            return cell.CellComment;
        }

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