//********************************************************************************************
//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

//works  
namespace Cliver
{
    public class Excel : IDisposable
    {
        static Excel()
        {
        }

        public Excel(string file, int worksheetId = 1)
        {
            File = file;
            init();
            OpenWorksheet(worksheetId);
        }

        public Excel(string file, string worksheetName)
        {
            File = file;
            init();
            OpenWorksheet(worksheetName);
        }

        void init()
        {
            if (System.IO.File.Exists(File))
                using (FileStream fs = new FileStream(File, FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        fs.Position = 0;//!!!prevents occasional error: EOF in header
                        workbook = new XSSFWorkbook(fs);
                    }
                    catch (ICSharpCode.SharpZipLib.Zip.ZipException)
                    {
                        fs.Position = 0;//!!!prevents occasional error: EOF in header
                        workbook = new HSSFWorkbook(fs);
                    }
                }
            else
                workbook = new XSSFWorkbook();
        }

        IWorkbook workbook;

        public readonly string File;

        ~Excel()
        {
            Dispose();
        }

        public void Dispose()
        {
            lock (this)
            {
                if (workbook != null)
                {
                    workbook.Close();
                    workbook = null;
                }
            }
        }

        public string HyperlinkBase
        {
            get
            {
                XSSFWorkbook xSSFWorkbook = workbook as XSSFWorkbook;
                if (xSSFWorkbook == null)
                    throw new Exception("TBD");
                NPOI.OpenXmlFormats.CT_Property p = xSSFWorkbook.GetProperties().CustomProperties.GetProperty("HyperlinkBase");//so is in Epplus
                return p?.Item?.ToString();
            }
            set
            {
                XSSFWorkbook xSSFWorkbook = workbook as XSSFWorkbook;
                if (xSSFWorkbook == null)
                    throw new Exception("TBD");
                List<NPOI.OpenXmlFormats.CT_Property> ps = xSSFWorkbook.GetProperties().CustomProperties.GetUnderlyingProperties().property;
                NPOI.OpenXmlFormats.CT_Property p = ps.Find(a => a.name == "HyperlinkBase");//so is in Epplus
                if (value == null)
                {
                    if (p != null)
                        ps.Remove(p);
                    return;
                }
                if (p == null)
                    xSSFWorkbook.GetProperties().CustomProperties.AddProperty("HyperlinkBase", value);
                else
                    p.Item = value;
            }
        }

        public void OpenWorksheet(string name)
        {
            worksheet = workbook.GetSheet(name);
            if (worksheet == null)
            {
                name = Regex.Replace(name, @"\:", "-");//npoi does not accept :
                worksheet = workbook.CreateSheet(WorkbookUtil.CreateSafeSheetName(name));

                //!!!All worksheet must be formatted as text! Otherwise string dates are converted into numbers.
                //!!!No way found to set default style for a whole sheet. However, NPOI presets ' before numeric values to keep them as strings.
                //ICellStyle defaultStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                //defaultStyle.DataFormat = workbook.CreateDataFormat().GetFormat("text");
                //ICell c = getCell(0, 0, true);
                //worksheet.SetDefaultColumnStyle(0, defaultStyle);
            }
        }
        //ICellStyle defaultStyle;

        public bool OpenWorksheet(int index)
        {
            index--;
            if (workbook.NumberOfSheets > 0 && workbook.NumberOfSheets > index)
            {
                worksheet = workbook.GetSheetAt(index);
                return true;
            }
            return false;
        }
        ISheet worksheet; //NPOI.XSSF.UserModel.XSSFSheet s;

        public string WorksheetName
        {
            get
            {
                return worksheet.SheetName;
            }
            set
            {
                if (worksheet != null)
                    workbook.SetSheetName(workbook.GetSheetIndex(worksheet), value);
            }
        }

        public void Save()
        {
            using (var fileData = new FileStream(File, FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }

        public int GetLastUsedRow()
        {
            if (worksheet == null)
                throw new Exception("No active sheet.");

            var rows = worksheet.GetRowEnumerator();
            int lur = 0;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (null != row.Cells.Find(a => !string.IsNullOrEmpty(a.ToString())))
                    lur = row.RowNum + 1;
            }
            return lur;
        }

        public int AppendLine(IEnumerable<object> values)
        {
            int y = GetLastUsedRow() + 1;
            int i = 1;
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;

                this[y, i++] = s;
            }
            return y;
        }

        public void SetLink(int y, int x, Uri uri)
        {
            int y0 = y - 1;
            int x0 = x - 1;
            ICell c = getCell(y0, x0, true);
            string v = c.ToString();
            if (string.IsNullOrEmpty(v))
                c.SetCellValue(LinkEmptyValueFiller);

            c.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
        }
        public static string LinkEmptyValueFiller = "           ";

        public Uri GetLink(int y, int x)
        {
            int y0 = y - 1;
            int x0 = x - 1;
            ICell c = getCell(y0, x0, false);
            if (c == null)
                return null;
            if (c.Hyperlink == null)
                return null;
            return new Uri(c.Hyperlink.Address, UriKind.RelativeOrAbsolute);
        }

        public string this[int y, int x]
        {
            get
            {
                int y0 = y - 1;
                int x0 = x - 1;
                ICell c = getCell(y0, x0, false);

            GET_VALUE: if (c == null)
                    return null;
                switch (c.CellType)
                {
                    case (CellType.Unknown):
                        return c.ToString();
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(c))
                        {
                            try
                            {
                                return c.DateCellValue.ToString("yyyy-MM-dd hh:mm:ss");
                            }
                            catch (Exception e)//!!!bug in NPOI2.5.1: after called Save(), it throw here NullReferenceException: GetLocaleCalendar()  https://github.com/nissl-lab/npoi/issues/358
                            {
                                Log.Warning("NPOI bug", e);
                                return DateTime.FromOADate(c.NumericCellValue).ToString("yyyy-MM-dd hh:mm:ss");
                            }
                            //return formatter.FormatCellValue(c);
                        }
                        return c.NumericCellValue.ToString();
                    case CellType.String:
                        return c.StringCellValue;
                    case CellType.Boolean:
                        return c.BooleanCellValue ? "TRUE" : "FALSE";
                    case CellType.Formula:
                        if (evaluator == null)
                            evaluator = new XSSFFormulaEvaluator(workbook);
                        c = evaluator.EvaluateInCell(c);
                        goto GET_VALUE;
                    //        return c.CellFormula;
                    case CellType.Error:
                        //return c.ErrorCellValue.ToString();
                        return FormulaError.ForInt(c.ErrorCellValue).String;
                    case CellType.Blank:
                        return string.Empty;
                    default:
                        throw new Exception("Unknown type: " + c.CellType);
                }
                //return getCell(y, x, false)?.ToString();
            }
            set
            {
                int y0 = y - 1;
                int x0 = x - 1;
                getCell(y0, x0, true).SetCellValue(value);
            }
        }
        IFormulaEvaluator evaluator = null;

        IRow getRow(int y0, bool create)
        {
            IRow r = worksheet.GetRow(y0);
            if (r == null && create)
            {
                r = worksheet.CreateRow(y0);
                ICellStyle cs = workbook.CreateCellStyle();
                cs.DataFormat = workbook.CreateDataFormat().GetFormat("text");
                r.RowStyle = cs;//!!!Cells must be formatted as text! Otherwise string dates are converted into numbers. (However, if no format set, NPOI presets ' before numeric values to keep them as strings.)
            }
            return r;
        }

        ICell getCell(int y0, int x0, bool create)
        {
            IRow r = getRow(y0, create);
            if (r == null)
                return null;
            return getCell(r, x0, create);
        }

        ICell getCell(IRow r, int x0, bool create)
        {
            ICell c = r.GetCell(x0);
            if (c != null)
                return c;
            if (create)
                return r.CreateCell(x0);
            return null;
        }

        public void InsertLine(int y, IEnumerable<object> values = null)
        {
            int y0 = y - 1;
            if (y <= worksheet.LastRowNum)
                worksheet.ShiftRows(y0, worksheet.LastRowNum, 1);
            getRow(y0, true);
            if (values != null)
                WriteLine(y, values);
        }

        public void WriteLine(int y, IEnumerable<object> values)
        {
            int y0 = y - 1;
            IRow r = getRow(y0, true);

            int x = 0;
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;

                getCell(r, x++, true).SetCellValue(s);
            }
        }

        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
        {
            int y0 = y - 1;
            int x0 = x - 1;
            List<string> vs = new List<string>();
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;
                vs.Add(s);
            }
            IDataValidationHelper dvh = new XSSFDataValidationHelper((XSSFSheet)worksheet);
            //string dvs = string.Join(",", vs);
            //IDataValidationConstraint dvc = worksheet.GetDataValidations().Find(a => string.Join(",", a.ValidationConstraint.ExplicitListValues) == dvs)?.ValidationConstraint;
            //if (dvc == null)
            //dvc = dvh.CreateCustomConstraint(dvs);
            IDataValidationConstraint dvc = dvh.CreateExplicitListConstraint(vs.ToArray());
            CellRangeAddressList cral = new CellRangeAddressList(y0, y0, x0, x0);
            IDataValidation dv = dvh.CreateValidation(dvc, cral);
            dv.SuppressDropDownArrow = true;
            dv.EmptyCellAllowed = allowBlank;
            ((XSSFSheet)worksheet).AddValidationData(dv);

            {
                string s;
                if (value is string)
                    s = (string)value;
                else if (value != null)
                    s = value.ToString();
                else
                    s = null;
                getCell(y0, x0, true).SetCellValue(s);
            }
        }

        public void AddImage(int y, int x, /*string name,*/ Bitmap image)//!!!!buggy
        {
            throw new Exception("TBD");
            int y0 = y - 1;
            int x0 = x - 1;
            int i = workbook.AddPicture(ImageToPngByteArray(image), PictureType.PNG);
            ICreationHelper h = workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = x0;//0 index based column
            a.Row1 = y0;//0 index based row
            XSSFDrawing d = (XSSFDrawing)worksheet.CreateDrawingPatriarch();
            XSSFPicture p = (XSSFPicture)d.CreatePicture(a, i);
            p.IsNoFill = true;
            p.Resize();
        }

        public static byte[] ImageToPngByteArray(Image img)
        {
            using (var stream = new MemoryStream())
            {
                img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }

        public Bitmap GetImage(int y, int x/*, out string name*/)//!!!!buggy
        {
            throw new Exception("TBD");
            //name = null;

            int y0 = y - 1;
            int x0 = x - 1;

            XSSFDrawing d = worksheet.CreateDrawingPatriarch() as XSSFDrawing;
            foreach (XSSFShape s in d.GetShapes())
            {
                XSSFPicture p = s as XSSFPicture;
                if (p == null)
                    continue;
                IClientAnchor a = p.GetPreferredSize();
                if (y0 >= a.Row1 && y0 <= a.Row2 && x0 >= a.Col1 && x0 <= a.Col2)
                {
                    XSSFPictureData pd = p.PictureData as XSSFPictureData;
                    //String ext = pd.SuggestFileExtension();
                    //name = pd.GetPackagePart().PartName?.ToString();
                    using (Stream ms = new MemoryStream(pd.Data))
                        return new Bitmap(ms);
                }
            }
            //name = null;
            return null;

            //var lst = workbook.GetAllPictures();
            //for (int i = 0; i < lst.Count; i++)
            //{
            //    var pd = (XSSFPictureData)lst[i];
            //    pd.RelationParts.Add[]
            //    using (Stream s = new MemoryStream(pd.Data))
            //        return new Bitmap(s);
            //}

            //foreach (NPOI.POIXMLDocumentPart dp in workbook.GetRelations())
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
            //            //int row2 = to.getRow();
            //            //int col2 = to.getCol();

            //            // do something here
            //        }
            //    }
            //}



            //foreach (XSSFPictureData pd in workbook.GetAllPictures())
            //{
            //    NPOI.OpenXml4Net.OPC.PackagePart pp = pd.GetPackagePart();
            //    pp.GetInputStream
            //  }
            //foreach (NPOI.POIXMLDocumentPart dp in workbook.GetRelations())
            //{
            //    NPOI.OpenXml4Net.OPC.PackagePart pp = dp.GetPackagePart();
            //    pp.GetInputStream
            //  }
        }

        public void FitColumnsWidth(IEnumerable<int> columnIs)
        {
            foreach (int i in columnIs)
                worksheet.AutoSizeColumn(i - 1);
        }

        public void FitColumnsWidth(int column1I, int column2I)
        {
            for (int i = column1I - 1; i < column2I; i++)
                worksheet.AutoSizeColumn(i);
        }

        public void HighlightRow(int y, Color color)
        {
            int y0 = y - 1;
            IRow r = getRow(y0, true);
            XSSFCellStyle cs = (XSSFCellStyle)r.RowStyle;
            if (cs == null)
            {
                cs = (XSSFCellStyle)workbook.CreateCellStyle();
                r.RowStyle = cs;
            }
            cs.SetFillForegroundColor(new XSSFColor(color));
            cs.FillPattern = FillPattern.SolidForeground;
        }

        public void ClearHighlighting()
        {
            var rows = worksheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                IRow r = (IRow)rows.Current;
                XSSFCellStyle cs = (XSSFCellStyle)r.RowStyle;
                if (cs != null)
                {
                    cs.FillPattern = FillPattern.NoFill;
                    cs.SetFillForegroundColor(null);
                }
                foreach (ICell c in r.Cells)
                {
                    cs = (XSSFCellStyle)c.CellStyle;
                    if (cs != null)
                    {
                        cs.FillPattern = FillPattern.NoFill;
                        cs.SetFillForegroundColor(null);
                    }
                }
            }
        }
    }
}