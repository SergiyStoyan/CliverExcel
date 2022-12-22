////********************************************************************************************
////Author: Sergiy Stoyan
////        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
////        http://www.cliversoft.com
////********************************************************************************************
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.IO;
//using System.Text.RegularExpressions;
//using System.Drawing;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml.Packaging;
//using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
//using A = DocumentFormat.OpenXml.Drawing;
//using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

//!!!To be developed!!!
//namespace Cliver.HawkeyeInvoiceParser
//{
//    public class Excel : IDisposable
//    {
//        static Excel()
//        {
//        }

//        public Excel(string file)
//        {
//            File = file;
//            init();
//            OpenWorksheet(0);
//        }

//        public Excel(string file, string worksheetName)
//        {
//            File = file;
//            init();
//            OpenWorksheet(worksheetName);
//        }

//        void init()
//        {
//            document = SpreadsheetDocument.Open(File, true);
//            //document.AutoSave = false;
//            workbookPart = document.WorkbookPart;
//        }

//        SpreadsheetDocument document;
//        WorkbookPart workbookPart;

//        public readonly string File;

//        ~Excel()
//        {
//            Dispose();
//        }

//        public void Dispose()
//        {
//            lock (this)
//            {
//                if (document != null)
//                {
//                    document.Close();
//                    document.Dispose();
//                    document = null;
//                }
//            }
//        }

//        public string HyperlinkBase
//        {
//            get
//            {
//                return document.ExtendedFilePropertiesPart.Properties.HyperlinkBase.Text;
//            }
//            set
//            {
//                document.ExtendedFilePropertiesPart.Properties.HyperlinkBase = new DocumentFormat.OpenXml.ExtendedProperties.HyperlinkBase(value);
//            }
//        }

//        public void OpenWorksheet(string name)
//        {
//            Workbook workbook = workbookPart.Workbook;
//            sheet = workbook.Descendants<Sheet>().Where(a => a.Name == name).FirstOrDefault();
//            if (sheet == null)
//                addBlankWorksheet(name);
//            worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
//        }

//        void addBlankWorksheet(string name)
//        {
//            WorksheetPart newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
//            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

//            Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
//            string relationshipId = document.WorkbookPart.GetIdOfPart(newWorksheetPart);

//            uint sheetId = 1;
//            if (sheets.Elements<Sheet>().Count() > 0)
//                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
//            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = name };
//            sheets.Append(sheet);
//        }

//        public bool OpenWorksheet(int index)
//        {
//            Workbook workbook = workbookPart.Workbook;
//            sheet = workbook.Descendants<Sheet>().Where(a => a.SheetId == index).FirstOrDefault();
//            if (sheet != null)
//                worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
//            return sheet != null;
//        }
//        Sheet sheet;
//        WorksheetPart worksheetPart;

//        public void Save()
//        {
//            document.WorkbookPart.Workbook.Save();
//            document.Save();
//        }

//        public string WorksheetName
//        {
//            get
//            {                
//                return sheet.Name;
//            }
//            set
//            {
//                if (sheet != null)
//                    sheet.Name=value;
//            }
//        }

//        public int GetLastUsedRow()
//        {
//            if (sheet == null)
//                throw new Exception("No active sheet.");
//            Row row = sheet.Descendants<Row>().LastOrDefault();
//            if (row != null)
//                return (int)row.RowIndex.Value;
//            return 0;
//        }

//        public int AppendLine(IEnumerable<object> values)
//        {
//            int y = GetLastUsedRow() + 1;
//            int i = 1;
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;

//                this[y, i++] = s;
//            }
//            return y;
//        }

//        public void SetLink(int y, int x, Uri uri)
//        {
//            Cell c = worksheetPart.Worksheet.Descendants<Cell>().Where(a => a.CellReference == GetCellReference(y, x)).FirstOrDefault();
//            if (c == null)
//                c = new Cell() { CellReference = GetCellReference(y, x), StyleIndex = (UInt32Value)1U, DataType = CellValues.InlineString };

//            string v = c.InnerText;
//            if (string.IsNullOrEmpty(v))
//                v = LinkEmptyValueFiller;

//            CellFormula cf = new CellFormula() { Space = SpaceProcessingModeValues.Preserve };
//            cf.Text = "HYPERLINK(\"" + uri.ToString() + "\", \"" + v + "\")";
//            c.RemoveAllChildren<CellFormula>();
//            c.Append(cf);

//            CellValue cv = new CellValue();
//            cv.Text = v;
//            c.RemoveAllChildren<CellValue>();
//            c.Append(cv);
//        }
//        public static string LinkEmptyValueFiller = "           ";

//        public static StringValue GetCellReference(int column, int row) =>
//        new StringValue($"{GetColumnName("", column)}{row}");

//        public static string GetColumnName(string prefix, int column) =>
//        column < 26 ? $"{prefix}{(char)(65 + column)}" :
//        GetColumnName(GetColumnName(prefix, (column - column % 26) / 26 - 1), column % 26);

//        public Uri GetLink(int y, int x)
//        {
//            Cell c = worksheetPart.Worksheet.Descendants<Cell>().Where(a => a.CellReference == GetCellReference(y, x)).FirstOrDefault();
//            if (c == null)
//                return null;
//            var hyperlinks = worksheetPart.RootElement.Descendants<Hyperlinks>().First().Cast<Hyperlink>();
//            var hyperlink = hyperlinks.SingleOrDefault(i => i.Reference.Value == c.CellReference.Value);
//            if (hyperlink == null)
//                return null;
//            var hyperlinksRelation = worksheetPart.HyperlinkRelationships.SingleOrDefault(i => i.Id == hyperlink.Id);
//            if (hyperlinksRelation == null)
//                return null;
//            return hyperlinksRelation.Uri;
//        }

//        public string this[int y, int x]
//        {
//            get
//            {
//                Cell c = worksheetPart.Worksheet.Descendants<Cell>().Where(a => a.CellReference == GetCellReference(y, x)).FirstOrDefault();
//                if (c == null)
//                    return null;
//                string value = c.InnerText;
//                // If the cell represents a numeric value, you are done. 
//                // For dates, this code returns the serialized value that 
//                // represents the date. The code handles strings and Booleans
//                // individually. For shared strings, the code looks up the 
//                // corresponding value in the shared string table. For Booleans, 
//                // the code converts the value into the words TRUE or FALSE.
//                if (c.DataType != null)
//                {
//                    switch (c.DataType.Value)
//                    {
//                        case CellValues.SharedString:
//                            // For shared strings, look up the value in the shared 
//                            // strings table.
//                            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
//                            // If the shared string table is missing, something is 
//                            // wrong. Return the index that you found in the cell.
//                            // Otherwise, look up the correct text in the table.
//                            if (stringTable != null)
//                            {
//                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
//                            }
//                            break;
//                        case CellValues.Boolean:
//                            switch (value)
//                            {
//                                case "0":
//                                    value = "FALSE";
//                                    break;
//                                default:
//                                    value = "TRUE";
//                                    break;
//                            }
//                            break;
//                    }
//                }
//                return value;
//            }
//            set
//            {
//                Cell c = worksheetPart.Worksheet.Descendants<Cell>().Where(a => a.CellReference == GetCellReference(y, x)).FirstOrDefault();
//                if (c == null)
//                    c = insertCell(y, x, worksheetPart.Worksheet);
//                c.CellValue = new CellValue(value);
//                //c.DataType = new EnumValue<CellValues>(CellValues.String);
//            }
//        }

//        private Cell insertCell(int rowIndex, int columnIndex, Worksheet worksheet)
//        {
//            Row row = null;
//            var sheetData = worksheet.GetFirstChild<SheetData>();

//            // Check if the worksheet contains a row with the specified row index.
//            row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
//            if (row == null)
//            {
//                row = new Row() { RowIndex = (uint)rowIndex };
//                sheetData.Append(row);
//            }

//            var cellReference = GetCellReference(columnIndex, rowIndex);      // e.g. A1

//            // Check if the row contains a cell with the specified column name.
//            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
//            if (cell == null)
//            {
//                cell = new Cell() { CellReference = cellReference };
//                if (row.ChildElements.Count < columnIndex)
//                    row.AppendChild(cell);
//                else
//                    row.InsertAt(cell, (int)columnIndex);
//            }

//            return cell;
//        }

//        public void InsertLine(int y, IEnumerable<object> values = null)
//        {
//            //worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(new Row() { RowIndex = y });
//            worksheetPart.Worksheet.GetFirstChild<SheetData>().InsertAt(new Row(), y);
//            sheet.InsertAt(new Row(), y);
//            if (values != null)
//                WriteLine(y, values);
//        }

//        public void WriteLine(int y, IEnumerable<object> values)
//        {
//            int i = 1;
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;

//                this[y, i++] = s;
//            }
//        }

//        public void CreateDropdown(int y, int x, IEnumerable<object> values, object value, bool allowBlank = true)
//        {
//            List<string> vs = new List<string>();
//            foreach (object v in values)
//            {
//                string s;
//                if (v is string)
//                    s = (string)v;
//                else if (v != null)
//                    s = v.ToString();
//                else
//                    s = null;
//                vs.Add(s);
//            }

//            DataValidation dv = new DataValidation
//            {
//                Type = DataValidationValues.List,
//                AllowBlank = allowBlank,
//                SequenceOfReferences = new ListValue<StringValue> { InnerText = GetCellReference(y, x) }
//            };

//            DataValidations dvs = worksheetPart.Worksheet.GetFirstChild<DataValidations>();
//            if (dvs == null)
//            {
//                dvs = new DataValidations();
//                dvs.Append(dv);
//                dvs.Count = 1;
//                worksheetPart.Worksheet.AppendChild(dvs);
//            }
//            else
//            {
//                dvs.Count = dvs.Count + 1;
//                dvs.Append(dv);
//            }

//            {
//                string s;
//                if (value is string)
//                    s = (string)value;
//                else if (value != null)
//                    s = value.ToString();
//                else
//                    s = null;
//                this[y, x] = s;
//            }
//        }

//        public void AddImage(int y, int x, string name, Bitmap image)
//        {
//        }

//        public Bitmap GetImage(int y, int x)
//        {
//            return null;
//        }

//        public void FitColumnsWidth(params int[] columnIs)
//        {
//            //foreach (int i in columnIs)
//            //    worksheetPart.Worksheet.Column(i).AutoFit();
//        }

//        public void HighlightRow(int y, System.Drawing.Color color)
//        {
//            //if (sheet.Row(y).Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.None)
//            //    sheet.Row(y).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
//            //sheet.Row(y).Style.Fill.BackgroundColor.SetColor(color);
//        }

//        public void ClearHighlighting()
//        {

//        }

















//        private void InsertImage(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
//        {
//            //Inserting a drawing element in worksheet
//            //Make sure that the relationship id is same for drawing element in worksheet and its relationship part
//            Drawing drawing1 = new Drawing() { Id = "rId1" };
//            worksheetPart.Worksheet.Append(drawing1);
//            //Adding the drawings.xml part
//            DrawingsPart drawingsPart1 = worksheetPart.AddNewPart<DrawingsPart>("rId1");
//            GenerateDrawingsPart1Content(drawingsPart1, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
//            //Adding the image
//            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
//            imagePart1.FeedData(imageStream);
//        }

//        // Generates content of drawingsPart1.
//        private static void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
//        {
//            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
//            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
//            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

//            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

//            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
//            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
//            columnId1.Text = startColumnIndex.ToString();
//            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
//            columnOffset1.Text = "38100";
//            Xdr.RowId rowId1 = new Xdr.RowId();
//            rowId1.Text = startRowIndex.ToString();
//            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
//            rowOffset1.Text = "0";

//            fromMarker1.Append(columnId1);
//            fromMarker1.Append(columnOffset1);
//            fromMarker1.Append(rowId1);
//            fromMarker1.Append(rowOffset1);

//            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
//            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
//            columnId2.Text = endColumnIndex.ToString();
//            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
//            columnOffset2.Text = "542925";
//            Xdr.RowId rowId2 = new Xdr.RowId();
//            rowId2.Text = endRowIndex.ToString();
//            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
//            rowOffset2.Text = "161925";

//            toMarker1.Append(columnId2);
//            toMarker1.Append(columnOffset2);
//            toMarker1.Append(rowId2);
//            toMarker1.Append(rowOffset2);

//            Xdr.Picture picture1 = new Xdr.Picture();

//            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
//            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

//            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
//            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

//            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

//            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
//            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

//            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

//            A.Blip blip1 = new A.Blip() { Embed = "rId1", CompressionState = A.BlipCompressionValues.Print };
//            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

//            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

//            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

//            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
//            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

//            blipExtension1.Append(useLocalDpi1);

//            blipExtensionList1.Append(blipExtension1);

//            blip1.Append(blipExtensionList1);

//            A.Stretch stretch1 = new A.Stretch();
//            A.FillRectangle fillRectangle1 = new A.FillRectangle();

//            stretch1.Append(fillRectangle1);

//            blipFill1.Append(blip1);
//            blipFill1.Append(stretch1);

//            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

//            A.Transform2D transform2D1 = new A.Transform2D();
//            A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
//            A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

//            transform2D1.Append(offset1);
//            transform2D1.Append(extents1);

//            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
//            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

//            presetGeometry1.Append(adjustValueList1);

//            shapeProperties1.Append(transform2D1);
//            shapeProperties1.Append(presetGeometry1);

//            picture1.Append(nonVisualPictureProperties1);
//            picture1.Append(blipFill1);
//            picture1.Append(shapeProperties1);
//            Xdr.ClientData clientData1 = new Xdr.ClientData();

//            twoCellAnchor1.Append(fromMarker1);
//            twoCellAnchor1.Append(toMarker1);
//            twoCellAnchor1.Append(picture1);
//            twoCellAnchor1.Append(clientData1);

//            worksheetDrawing1.Append(twoCellAnchor1);

//            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
//        }






























//        /// <summary>
//        /// Inserts a new row at the desired index. If one already exists, then it is
//        /// returned. If an insertRow is provided, then it is inserted into the desired
//        /// rowIndex
//        /// </summary>
//        /// <param name="rowIndex">Row Index</param>
//        /// <param name="worksheetPart">Worksheet Part</param>
//        /// <param name="insertRow">Row to insert</param>
//        /// <param name="isLastRow">Optional parameter - True, you can guarantee that this row is the last row (not replacing an existing last row) in the sheet to insert; false it is not</param>
//        /// <returns>Inserted Row</returns>
//        Row insertRow(uint rowIndex, Row insertRow, bool isNewLastRow = false)
//        {
//            Worksheet worksheet = worksheetPart.Worksheet;
//            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

//            Row retRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex) : null;

//            // If the worksheet does not contain a row with the specified row index, insert one.
//            if (retRow != null)
//            {
//                // if retRow is not null and we are inserting a new row, then move all existing rows down.
//                if (insertRow != null)
//                {
//                    UpdateRowIndexes(worksheetPart, rowIndex, false);
//                    UpdateMergedCellReferences(worksheetPart, rowIndex, false);
//                    UpdateHyperlinkReferences(worksheetPart, rowIndex, false);

//                    // actually insert the new row into the sheet
//                    retRow = sheetData.InsertBefore(insertRow, retRow);  // at this point, retRow still points to the row that had the insert rowIndex

//                    string curIndex = retRow.RowIndex.ToString();
//                    string newIndex = rowIndex.ToString();

//                    foreach (Cell cell in retRow.Elements<Cell>())
//                    {
//                        // Update the references for the rows cells.
//                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
//                    }

//                    // Update the row index.
//                    retRow.RowIndex = rowIndex;
//                }
//            }
//            else
//            {
//                // Row doesn't exist yet, shifting not needed.
//                // Rows must be in sequential order according to RowIndex. Determine where to insert the new row.
//                Row refRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex > rowIndex) : null;

//                // use the insert row if it exists
//                retRow = insertRow ?? new Row() { RowIndex = rowIndex };

//                IEnumerable<Cell> cellsInRow = retRow.Elements<Cell>();

//                if (cellsInRow.Any())
//                {
//                    string curIndex = retRow.RowIndex.ToString();
//                    string newIndex = rowIndex.ToString();

//                    foreach (Cell cell in cellsInRow)
//                    {
//                        // Update the references for the rows cells.
//                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
//                    }

//                    // Update the row index.
//                    retRow.RowIndex = rowIndex;
//                }

//                sheetData.InsertBefore(retRow, refRow);
//            }

//            return retRow;
//        }
//        /// <summary>
//        /// Updates all of the Row indexes and the child Cells' CellReferences whenever
//        /// a row is inserted or deleted.
//        /// </summary>
//        /// <param name="worksheetPart">Worksheet Part</param>
//        /// <param name="rowIndex">Row Index being inserted or deleted</param>
//        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
//        private static void UpdateRowIndexes(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
//        {
//            // Get all the rows in the worksheet with equal or higher row index values than the one being inserted/deleted for reindexing.
//            IEnumerable<Row> rows = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

//            foreach (Row row in rows)
//            {
//                uint newIndex = (isDeletedRow ? row.RowIndex - 1 : row.RowIndex + 1);
//                string curRowIndex = row.RowIndex.ToString();
//                string newRowIndex = newIndex.ToString();

//                foreach (Cell cell in row.Elements<Cell>())
//                {
//                    // Update the references for the rows cells.
//                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex));
//                }

//                // Update the row index.
//                row.RowIndex = newIndex;
//            }
//        }

//        /// <summary>
//        /// Updates the MergedCelss reference whenever a new row is inserted or deleted. It will simply take the
//        /// row index and either increment or decrement the cell row index in the merged cell reference based on
//        /// if the row was inserted or deleted.
//        /// </summary>
//        /// <param name="worksheetPart">Worksheet Part</param>
//        /// <param name="rowIndex">Row Index being inserted or deleted</param>
//        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
//        private static void UpdateMergedCellReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
//        {
//            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
//            {
//                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

//                if (mergeCells != null)
//                {
//                    // Grab all the merged cells that have a merge cell row index reference equal to or greater than the row index passed in
//                    List<MergeCell> mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
//                                                                                     .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) >= rowIndex ||
//                                                                                                 GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) >= rowIndex).ToList();

//                    // Need to remove all merged cells that have a matching rowIndex when the row is deleted
//                    if (isDeletedRow)
//                    {
//                        List<MergeCell> mergeCellsToDelete = mergeCellsList.Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) == rowIndex ||
//                                                                                       GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) == rowIndex).ToList();

//                        // Delete all the matching merged cells
//                        foreach (MergeCell cellToDelete in mergeCellsToDelete)
//                        {
//                            cellToDelete.Remove();
//                        }

//                        // Update the list to contain all merged cells greater than the deleted row index
//                        mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
//                                                                         .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) > rowIndex ||
//                                                                                     GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) > rowIndex).ToList();
//                    }

//                    // Either increment or decrement the row index on the merged cell reference
//                    foreach (MergeCell mergeCell in mergeCellsList)
//                    {
//                        string[] cellReference = mergeCell.Reference.Value.Split(':');

//                        if (GetRowIndex(cellReference.ElementAt(0)) >= rowIndex)
//                        {
//                            string columnName = GetColumnName(cellReference.ElementAt(0));
//                            cellReference[0] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(0)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(0), CellReferencePartEnum.Row);
//                        }

//                        if (GetRowIndex(cellReference.ElementAt(1)) >= rowIndex)
//                        {
//                            string columnName = GetColumnName(cellReference.ElementAt(1));
//                            cellReference[1] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(1)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(1), CellReferencePartEnum.Row);
//                        }

//                        mergeCell.Reference = new StringValue(cellReference[0] + ":" + cellReference[1]);
//                    }
//                }
//            }
//        }

//        /// <summary>
//        /// Updates all hyperlinks in the worksheet when a row is inserted or deleted.
//        /// </summary>
//        /// <param name="worksheetPart">Worksheet Part</param>
//        /// <param name="rowIndex">Row Index being inserted or deleted</param>
//        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
//        private static void UpdateHyperlinkReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
//        {
//            Hyperlinks hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();

//            if (hyperlinks != null)
//            {
//                Match hyperlinkRowIndexMatch;
//                uint hyperlinkRowIndex;

//                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>())
//                {
//                    hyperlinkRowIndexMatch = Regex.Match(hyperlink.Reference.Value, "[0-9]+");
//                    if (hyperlinkRowIndexMatch.Success && uint.TryParse(hyperlinkRowIndexMatch.Value, out hyperlinkRowIndex) && hyperlinkRowIndex >= rowIndex)
//                    {
//                        // if being deleted, hyperlink needs to be removed or moved up
//                        if (isDeletedRow)
//                        {
//                            // if hyperlink is on the row being removed, remove it
//                            if (hyperlinkRowIndex == rowIndex)
//                            {
//                                hyperlink.Remove();
//                            }
//                            // else hyperlink needs to be moved up a row
//                            else
//                            {
//                                hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex - 1).ToString());

//                            }
//                        }
//                        // else row is being inserted, move hyperlink down
//                        else
//                        {
//                            hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex + 1).ToString());
//                        }
//                    }
//                }

//                // Remove the hyperlinks collection if none remain
//                if (hyperlinks.Elements<Hyperlink>().Count() == 0)
//                {
//                    hyperlinks.Remove();
//                }
//            }
//        }

//        /// <summary>
//        /// Given a cell name, parses the specified cell to get the row index.
//        /// </summary>
//        /// <param name="cellReference">Address of the cell (ie. B2)</param>
//        /// <returns>Row Index (ie. 2)</returns>
//        public static uint GetRowIndex(string cellReference)
//        {
//            // Create a regular expression to match the row index portion the cell name.
//            Regex regex = new Regex(@"\d+");
//            Match match = regex.Match(cellReference);

//            return uint.Parse(match.Value);
//        }

//        /// <summary>
//        /// Increments the reference of a given cell.  This reference comes from the CellReference property
//        /// on a Cell.
//        /// </summary>
//        /// <param name="reference">reference string</param>
//        /// <param name="cellRefPart">indicates what is to be incremented</param>
//        /// <returns></returns>
//        public static string IncrementCellReference(string reference, CellReferencePartEnum cellRefPart)
//        {
//            string newReference = reference;

//            if (cellRefPart != CellReferencePartEnum.None && !String.IsNullOrEmpty(reference))
//            {
//                string[] parts = Regex.Split(reference, "([A-Z]+)");

//                if (cellRefPart == CellReferencePartEnum.Column || cellRefPart == CellReferencePartEnum.Both)
//                {
//                    List<char> col = parts[1].ToCharArray().ToList();
//                    bool needsIncrement = true;
//                    int index = col.Count - 1;

//                    do
//                    {
//                        // increment the last letter
//                        col[index] = Letters[Letters.IndexOf(col[index]) + 1];

//                        // if it is the last letter, then we need to roll it over to 'A'
//                        if (col[index] == Letters[Letters.Count - 1])
//                        {
//                            col[index] = Letters[0];
//                        }
//                        else
//                        {
//                            needsIncrement = false;
//                        }

//                    } while (needsIncrement && --index >= 0);

//                    // If true, then we need to add another letter to the mix. Initial value was something like "ZZ"
//                    if (needsIncrement)
//                    {
//                        col.Add(Letters[0]);
//                    }

//                    parts[1] = new String(col.ToArray());
//                }

//                if (cellRefPart == CellReferencePartEnum.Row || cellRefPart == CellReferencePartEnum.Both)
//                {
//                    // Increment the row number. A reference is invalid without this componenet, so we assume it will always be present.
//                    parts[2] = (int.Parse(parts[2]) + 1).ToString();
//                }

//                newReference = parts[1] + parts[2];
//            }

//            return newReference;
//        }

//        /// <summary>
//        /// Given a cell name, parses the specified cell to get the column name.
//        /// </summary>
//        /// <param name="cellReference">Address of the cell (ie. B2)</param>
//        /// <returns>Column name (ie. A2)</returns>
//        private static string GetColumnName(string cellName)
//        {
//            // Create a regular expression to match the column name portion of the cell name.
//            Regex regex = new Regex("[A-Za-z]+");
//            Match match = regex.Match(cellName);

//            return match.Value;
//        }
//        public enum CellReferencePartEnum
//        {
//            None,
//            Column,
//            Row,
//            Both
//        }
//        private static List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };
//    }
//}