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
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;

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
                        Workbook = new XSSFWorkbook(fs);
                        FormulaEvaluator = new XSSFFormulaEvaluator(Workbook);
                    }
                    catch (ICSharpCode.SharpZipLib.Zip.ZipException)
                    {
                        fs.Position = 0;//!!!prevents error: EOF in header
                        Workbook = new HSSFWorkbook(fs);//old Excel 97-2003
                        FormulaEvaluator = new HSSFFormulaEvaluator(Workbook);
                    }
                }
            else
                Workbook = new XSSFWorkbook();
        }

        public IWorkbook Workbook;
        public IFormulaEvaluator FormulaEvaluator = null;

        public readonly string File;

        ~Excel()
        {
            Dispose();
        }

        public void Dispose()
        {
            lock (this)
            {
                if (Workbook != null)
                {
                    Workbook.Close();
                    Workbook = null;
                }
            }
        }

        public string HyperlinkBase
        {
            get
            {
                XSSFWorkbook xSSFWorkbook = Workbook as XSSFWorkbook;
                if (xSSFWorkbook == null)
                    throw new Exception("TBD");
                NPOI.OpenXmlFormats.CT_Property p = xSSFWorkbook.GetProperties().CustomProperties.GetProperty("HyperlinkBase");//so is in Epplus
                return p?.Item?.ToString();
            }
            set
            {
                XSSFWorkbook xSSFWorkbook = Workbook as XSSFWorkbook;
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
            worksheet = Workbook.GetSheet(name);
            if (worksheet == null)
            {
                name = Regex.Replace(name, @"\:", "-");//npoi does not accept :
                worksheet = Workbook.CreateSheet(WorkbookUtil.CreateSafeSheetName(name));

                //!!!All worksheet must be formatted as text! Otherwise string dates are converted into numbers.
                //!!!No way found to set default style for a whole sheet. However, NPOI presets ' before numeric values to keep them as strings.
                //ICellStyle defaultStyle = (XSSFCellStyle)Workbook.CreateCellStyle();
                //defaultStyle.DataFormat = Workbook.CreateDataFormat().GetFormat("text");
                //ICell c = GetCell(0, 0, true);
                //worksheet.SetDefaultColumnStyle(0, defaultStyle);
            }
        }
        //ICellStyle defaultStyle;

        public bool OpenWorksheet(int index)
        {
            index--;
            if (Workbook.NumberOfSheets > 0 && Workbook.NumberOfSheets > index)
            {
                worksheet = Workbook.GetSheetAt(index);
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
                    Workbook.SetSheetName(Workbook.GetSheetIndex(worksheet), value);
            }
        }

        public void Save()
        {
            using (var fileData = new FileStream(File, FileMode.Create))
            {
                Workbook.Write(fileData);
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
                    lur = row.RowNum;
            }
            return lur + 1;
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
            ICell c = GetCell(y, x, true);
            if (string.IsNullOrEmpty(this[y, x]))
                c.SetCellValue(LinkEmptyValueFiller);
            if (Workbook is XSSFWorkbook)
                c.Hyperlink = new XSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
            else if (Workbook is HSSFWorkbook)
                c.Hyperlink = new HSSFHyperlink(HyperlinkType.Url) { Address = uri.ToString() };
        }
        public string LinkEmptyValueFiller = "           ";

        public Uri GetLink(int y, int x)
        {
            ICell c = GetCell(y, x, false);
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
                ICell c = GetCell(y, x, false);
                return ExcelExtensions.GetValueAsString(c, FormulaEvaluator);
            }
            set
            {
                ICell c = GetCell(y, x, true);
                //c.SetBlank();
                //c.SetCellType(CellType.String);
                c.SetCellValue(value);
            }
        }

        public IRow GetRow(int y, bool create)
        {
            IRow r = worksheet.GetRow(y - 1);
            if (r == null && create)
            {
                r = worksheet.CreateRow(y - 1);
                ICellStyle cs = Workbook.CreateCellStyle();
                cs.DataFormat = Workbook.CreateDataFormat().GetFormat("text");
                r.RowStyle = cs;//!!!Cells must be formatted as text! Otherwise string dates are converted into numbers. (However, if no format set, NPOI presets ' before numeric values to keep them as strings.)
            }
            return r;
        }

        public ICell GetCell(int y, int x, bool create)
        {
            IRow r = GetRow(y, create);
            if (r == null)
                return null;
            return r.GetCell(x, create);
        }

        public void InsertLine(int y, IEnumerable<object> values = null)
        {
            if (y <= worksheet.LastRowNum)
                worksheet.ShiftRows(y - 1, worksheet.LastRowNum, 1);
            GetRow(y, true);
            if (values != null)
                WriteLine(y, values);
        }

        public void WriteLine(int y, IEnumerable<object> values)
        {
            IRow r = GetRow(y, true);

            int x = 1;
            foreach (object v in values)
            {
                string s;
                if (v is string)
                    s = (string)v;
                else if (v != null)
                    s = v.ToString();
                else
                    s = null;

                r.GetCell(x++, true).SetCellValue(s);
            }
        }

        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
        {
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
            CellRangeAddressList cral = new CellRangeAddressList(y - 1, y - 1, x - 1, x - 1);
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
                GetCell(y, x, true).SetCellValue(s);
            }
        }

        public void AddImage(int y, int x, /*string name,*/ Bitmap image)//!!!!buggy
        {
            throw new Exception("TBD");
            int i = Workbook.AddPicture(ImageToPngByteArray(image), PictureType.PNG);
            ICreationHelper h = Workbook.GetCreationHelper();
            IClientAnchor a = h.CreateClientAnchor();
            a.AnchorType = AnchorType.MoveDontResize;
            a.Col1 = x - 1;//0 index based column
            a.Row1 = y - 1;//0 index based row
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

            XSSFDrawing d = worksheet.CreateDrawingPatriarch() as XSSFDrawing;
            foreach (XSSFShape s in d.GetShapes())
            {
                XSSFPicture p = s as XSSFPicture;
                if (p == null)
                    continue;
                IClientAnchor a = p.GetPreferredSize();
                if (y - 1 >= a.Row1 && y - 1 <= a.Row2 && x - 1 >= a.Col1 && x - 1 <= a.Col2)
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

            //var lst = Workbook.GetAllPictures();
            //for (int i = 0; i < lst.Count; i++)
            //{
            //    var pd = (XSSFPictureData)lst[i];
            //    pd.RelationParts.Add[]
            //    using (Stream s = new MemoryStream(pd.Data))
            //        return new Bitmap(s);
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



            //foreach (XSSFPictureData pd in Workbook.GetAllPictures())
            //{
            //    NPOI.OpenXml4Net.OPC.PackagePart pp = pd.GetPackagePart();
            //    pp.GetInputStream
            //  }
            //foreach (NPOI.POIXMLDocumentPart dp in Workbook.GetRelations())
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

        public void FitColumnsWidth(int x1, int x2)
        {
            for (int x = x1 - 1; x < x2; x++)
                worksheet.AutoSizeColumn(x);
        }

        public void HighlightRow(int y, Color color)
        {
            IRow r = GetRow(y, true);
            XSSFCellStyle cs = (XSSFCellStyle)r.RowStyle;
            if (cs == null)
            {
                cs = (XSSFCellStyle)Workbook.CreateCellStyle();
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

        public void CopyRange(CellRangeAddress range, ISheet sourceSheet, ISheet destinationSheet)
        {
            for (int y = range.FirstRow; y <= range.LastRow; y++)
            {
                IRow sourceRow = sourceSheet.GetRow(y);
                if (sourceRow == null)
                    continue;
                IRow destinationRow = destinationSheet.GetRow(y);
                if (destinationRow == null)
                    destinationRow = destinationSheet.CreateRow(y);
                for (int x = range.FirstColumn; x < sourceRow.LastCellNum && x <= range.LastColumn; x++)
                {
                    ICell sourceCell = sourceRow.GetCell(x);
                    ICell destinationCell = destinationRow.GetCell(x);
                    if (sourceCell == null)
                    {
                        if (destinationCell == null)
                            continue;
                        destinationRow.RemoveCell(destinationCell);
                    }
                    else
                    {
                        destinationCell = destinationRow.CreateCell(x);
                        CopyCell(sourceCell, destinationCell);
                    }
                }
            }
        }

        public void CopyColumn(string columnName, ISheet sourceSheet, ISheet destinationSheet)
        {
            int x = CellReference.ConvertColStringToIndex(columnName);
            CopyColumn(x, sourceSheet, destinationSheet);
        }

        public void CopyColumn(int x, ISheet sourceSheet, ISheet destinationSheet)
        {
            var range = new CellRangeAddress(0, sourceSheet.LastRowNum, x - 1, x - 1);
            CopyRange(range, sourceSheet, destinationSheet);
        }

        public void CopyCell(ICell source, ICell destination)
        {
            destination.SetBlank();
            destination.SetCellType(source.CellType);
            destination.CellStyle = source.CellStyle;
            destination.CellComment = source.CellComment;
            destination.Hyperlink = source.Hyperlink;
            switch (source.CellType)
            {
                case CellType.Formula:
                    destination.CellFormula = source.CellFormula;
                    break;
                case CellType.Numeric:
                    destination.SetCellValue(source.NumericCellValue);
                    break;
                case CellType.String:
                    destination.SetCellValue(source.StringCellValue);
                    break;
                case CellType.Boolean:
                    destination.SetCellValue(source.BooleanCellValue);
                    break;
                case CellType.Error:
                    destination.SetCellErrorValue(source.ErrorCellValue);
                    break;
                case CellType.Blank:
                    destination.SetBlank();
                    break;
                default:
                    throw new Exception("Unknown cell type: " + source.CellType);
            }
        }

        public void CopyCell(ICell sourceCell, int destinationY, int destinationX)
        {
            if (sourceCell == null)
            {
                IRow destinationRow = GetRow(destinationY, false);
                if (destinationRow == null)
                    return;
                ICell destinationCell = destinationRow.GetCell(destinationX, false);
                if (destinationCell == null)
                    return;
                destinationRow.RemoveCell(destinationCell);
            }
            else
            {
                ICell destinationCell = GetCell(destinationY, destinationX, true);
                CopyCell(sourceCell, destinationCell);
            }
        }

        public void CopyCell(int sourceY, int sourceX, int destinationY, int destinationX)
        {
            ICell sourceCell = GetCell(sourceY, sourceX, false);
            CopyCell(sourceCell, destinationY, destinationX);
        }

        public void CopyRange(Range sourceRange, Point destinationPoint)
        {
            ICell[,] sourceCells = new ICell[sourceRange.Bottom - sourceRange.Y + 1, sourceRange.Right - sourceRange.X + 1];
            for (int y = sourceRange.Y - 1; y < sourceRange.Bottom; y++)
            {
                IRow sourceRow = worksheet.GetRow(y);
                if (sourceRow == null)
                    continue;
                for (int x = sourceRange.X - 1; x <= sourceRow.LastCellNum && x < sourceRange.Right; x++)
                    sourceCells[y, x] = sourceRow.GetCell(x);
            }
            int height = sourceCells.GetLength(0);
            int width = sourceCells.GetLength(1);
            for (int y = 0; y < height; y++)
            {
                IRow destinationRow = worksheet.GetRow(y + (destinationPoint.Y - sourceRange.Y));
                for (int x = 0; x < width; x++)
                {
                    if (sourceCells[y, x] == null)
                    {
                        if (destinationRow == null)
                            continue;
                        ICell destinationCell = destinationRow.GetCell(x + (destinationPoint.X - sourceRange.X));
                        if (destinationCell == null)
                            continue;
                        destinationRow.RemoveCell(destinationCell);
                    }
                    else
                        CopyCell(sourceCells[y, x], y + (destinationPoint.Y - sourceRange.Y), x + (destinationPoint.X - sourceRange.X));
                }
            }
        }

        public class Range
        {
            public int X;
            public int Right;
            public int Y;
            public int Bottom;
        }

        public class Point
        {
            public int X;
            public int Y;
            //public ICell _Cell;
        }

        public int GetLastUsedRowInColumn(int x)
        {
            if (worksheet == null)
                throw new Exception("No active sheet.");

            var rows = worksheet.GetRowEnumerator();
            int lur = 0;
            int x0 = x - 1;
            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                if (!string.IsNullOrEmpty(row.GetCell(x0)?.ToString()))
                    lur = row.RowNum;
            }
            return lur + 1;
        }

        public void ShiftCellsDown(int cellsY, int firstCellX, int lastCellX, int rowCount, Action<ICell> updateFormula = null)
        {
            for (int x = firstCellX; x <= lastCellX; x++)
            {
                for (int y = GetLastUsedRowInColumn(x); y >= cellsY; y--)
                {
                    CopyCell(y, x, y + rowCount, x);
                    if (updateFormula == null)
                        continue;
                    ICell formulaCell = GetCell(y + rowCount, x, false);
                    if (formulaCell?.CellType != CellType.Formula)
                        continue;
                    updateFormula(formulaCell);
                }
                GetCell(cellsY, x, false)?.SetBlank();
            }
        }

        public void UpdateFormulaRange(ICell formulaCell, int rangeY1Shift, int rangeX1Shift, int? rangeY2Shift = null, int? rangeX2Shift = null)
        {
            if (rangeY2Shift == null)
                rangeY2Shift = rangeY1Shift;
            if (rangeX2Shift == null)
                rangeX2Shift = rangeX1Shift;

            IFormulaParsingWorkbook evaluationWorkbook;
            if (Workbook is XSSFWorkbook)
                evaluationWorkbook = XSSFEvaluationWorkbook.Create(Workbook);
            else if (Workbook is HSSFWorkbook)
                evaluationWorkbook = HSSFEvaluationWorkbook.Create(Workbook);
            //else if (sheet is SXSSFWorkbook)
            //{
            //    evaluationWorkbook = SXSSFEvaluationWorkbook.Create((SXSSFWorkbook)Workbook);
            else
                throw new Exception("Unexpected Workbook type: " + Workbook.GetType());

            //ICell formulaCell = GetCell(formulaCellY, formulaCellX, false);
            if (formulaCell?.CellType != CellType.Formula)
                return;
            var ptgs = FormulaParser.Parse(formulaCell.CellFormula, evaluationWorkbook, FormulaType.Cell, Workbook.GetSheetIndex(worksheet));
            foreach (Ptg ptg in ptgs)
            {
                if (ptg is RefPtgBase)
                {
                    RefPtgBase ref2 = (RefPtgBase)ptg;
                    if (ref2.IsRowRelative)
                        ref2.Row = ref2.Row + rangeY1Shift;
                    if (ref2.IsColRelative)
                        ref2.Column = ref2.Column + rangeX1Shift;
                }
                else if (ptg is AreaPtgBase)
                {
                    AreaPtgBase ref2 = (AreaPtgBase)ptg;
                    if (ref2.IsFirstRowRelative)
                        ref2.FirstRow += rangeY1Shift;
                    if (ref2.IsLastRowRelative)
                        ref2.LastRow += rangeY2Shift.Value;
                    if (ref2.IsFirstColRelative)
                        ref2.FirstColumn += rangeX1Shift;
                    if (ref2.IsLastColRelative)
                        ref2.LastColumn += rangeX2Shift.Value;
                }
                //else
                //    throw new Exception("Unexpected ptg type: " + ptg.GetType());
            }
            formulaCell.CellFormula = FormulaRenderer.ToFormulaString((IFormulaRenderingWorkbook)evaluationWorkbook, ptgs);
        }
    }
}