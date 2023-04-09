//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************

using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using System.Linq;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using static Cliver.Excel;

namespace Cliver
{
    public partial class Sheet
    {
        public Sheet(ISheet sheet)
        {
            _ = sheet;
        }
        public ISheet _ { get; private set; }

        public void ReplaceStyle(ICellStyle style1, ICellStyle style2)
        {
            new Excel.Range(_).ReplaceStyle(style1, style2);
        }

        public void SetStyle(ICellStyle style, bool createCells)
        {
            new Excel.Range(_).SetStyle(style, createCells);
        }

        public void UnsetStyle(ICellStyle style)
        {
            new Excel.Range(_).UnsetStyle(style);
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <exception cref="Exception"></exception>
        public void AddImage(Excel.Image image)
        {
            int i = _.Workbook.AddPicture(image.Data, image.Type);
            ICreationHelper h = _.Workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = image.X - 1;//0-based column index
            a.Row1 = image.Y - 1;//0-based row index
            IDrawing d = _.CreateDrawingPatriarch();
            IPicture p = d.CreatePicture(a, i);
            if (p is XSSFPicture xSSFPicture)
                xSSFPicture.IsNoFill = true;
            //p.Resize();
        }

        /// <summary>
        /// !!!BUGGY!!!
        /// </summary>
        /// <param name="y"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public IEnumerable<Excel.Image> GetImages(int y, int x)
        {
            foreach (IPicture p in _.Workbook.GetAllPictures())
            {
                if (p == null)
                    continue;
                IClientAnchor a = p.GetPreferredSize();
                if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
                {
                    IPictureData pictureData = p.PictureData;
                    yield return new Excel.Image { Data = pictureData.Data, Name = null, Type = pictureData.PictureType, X = a.Col1, Y = a.Row1/*, Anchor = a*/ };
                }
            }

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

        public Excel.Range NewRange(int y1 = 1, int x1 = 1, int? y2 = null, int? x2 = null)
        {
            return new Excel.Range(_, y1, x1, y2, x2);
        }

        public Excel.Range GetMergedRange(int y, int x)
        {
            return _.getMergedRange(y, x);
        }
    }
}