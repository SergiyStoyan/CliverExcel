//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************

using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using static Cliver.Excel;
using System.Linq;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace Cliver
{
    static public partial class ExcelExtensions
    {
        static public void _ReplaceStyle(this ISheet sheet, ICellStyle style1, ICellStyle style2)
        {
            new Range(sheet).ReplaceStyle(style1, style2);
        }

        static public void _SetStyle(this ISheet sheet, ICellStyle style, bool createCells)
        {
            new Range(sheet).SetStyle(style, createCells);
        }

        static public void _UnsetStyle(this ISheet sheet, ICellStyle style)
        {
            new Range(sheet).UnsetStyle(style);
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        static public IEnumerable<Image> _GetImages(this ISheet sheet, int y, int x)
        {
            if (sheet.Workbook is XSSFWorkbook xSSFWorkbook)
            {
                XSSFDrawing dp = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                foreach (XSSFShape s in dp.GetShapes())
                {
                    XSSFPicture p = s as XSSFPicture;
                    if (p == null)
                        continue;
                    var a = p.ClientAnchor;
                    if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
                    {
                        IPictureData pictureData = p.PictureData;
                        yield return new Image { Data = pictureData.Data, Name = null, Type = pictureData.PictureType, X = a.Col1, Y = a.Row1/*, Anchor = a*/ };
                    }
                }
            }
            else if (sheet.Workbook is HSSFWorkbook hSSFWorkbook)
            {
                HSSFPatriarch g;

            }
            else
                throw new Exception("Unsupported workbook type: " + sheet.Workbook.GetType().FullName);



            //XSSFDrawing d = Sheet.CreateDrawingPatriarch() as XSSFDrawing;
            //foreach (XSSFShape s in d.GetShapes())
            //{
            //    XSSFPicture p = s as XSSFPicture;
            //    if (p == null)
            //        continue;
            //    IClientAnchor a = p.GetPreferredSize();
            //    if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
            //    {
            //        XSSFPictureData pd = p.PictureData as XSSFPictureData;
            //        pictureType = pd.PictureType;
            //        return pd.Data;
            //    }
            //}

            //foreach (NPOI.POIXMLDocumentPart dp in Workbook.GetRelations())
            //{
            //    if (dp is XSSFDrawing)
            //    {
            //        NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing d = ((XSSFDrawing)dp).GetCTDrawing();
            //        foreach (NPOI.OpenXmlFormats.Dml.Spreadsheet.IEG_Anchor a in d.CellAnchors)
            //        {
            //            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor aa = a as NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor;
            //            if (aa == null)
            //                continue;
            //            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Marker m = aa.from;
            //            if (m.row == y && m.col == x)
            //                //using (Stream s = new MemoryStream(((XSSFPicture)aa.picture).PictureData))
            //                    return new Bitmap(0,0);
            //            //CTMarker to = anchor.getTo();
            //            //int row2 = to.GetRow();
            //            //int col2 = to.getCol();

            //            // do something here
            //        }
            //    }
            //}
        }

        static public Range _NewRange(this ISheet sheet, int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return new Range(sheet, y1, x1, y2, x2);
        }

        static public Range _GetMergedRange(this ISheet sheet, int y, int x)
        {
            return sheet._getMergedRange(y, x);
        }
    }
}