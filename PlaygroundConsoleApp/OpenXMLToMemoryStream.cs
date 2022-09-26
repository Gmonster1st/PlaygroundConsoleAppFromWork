using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;

namespace PlaygroundConsoleApp
{
    internal class OpenXMLToMemoryStream
    {
        // These are class level variables set in createstylesheet
        private uint CELLSTYLE_DEFAULT = 0;

        private uint HEADER_CELLSTYLE_LEFT_JUSTIFIED = 0;
        private uint HEADER_CELLSTYLE_RIGHT_JUSTIFIED = 0;
        private uint DATA_CELLSTYLE_TEXT = 0;
        private uint DATA_CELLSTYLE_DATE = 0;
        private uint DATA_CELLSTYLE_INVOICE_NUMBER = 0;
        private uint DATA_CELLSTYLE_WILL_PICKUP = 0;
        private uint DATA_CELLSTYLE_QTY = 0;
        private uint DATA_CELLSTYLE_CURRENCY = 0;
        private uint FOOTER_CELLSTYLE_TOTAL_CURRENCY = 0;
        private uint FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT = 0;

        #region CompleteWorksheet

        // ******************************************************************************
        // ******************************************************************************
        // This action Creates an Excel File In Memory and Returns it as an ActionResult
        // ******************************************************************************
        // ******************************************************************************
        public IActionResult CreateExcelFile()
        {
            IActionResult rslt = new BadRequestResult();

            string WorksheetName = "CompleteWorksheet";

            // create a new memory stream;
            MemoryStream ms = new MemoryStream();

            // ---------------------------- BEGIN CONSTRUCTING WORKBOOK -------------------

            // boilerplate
            // create a new document to the stream
            SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, false);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = doc.AddWorkbookPart();

            // add the workbook to the workbookpart (1 workbook for workbookpart)
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // add a new worksheet to the worksheet part.  Note the initialization
            // of SheetData to the worksheet contructor
            // worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet = new Worksheet();

            // ---------------------------- BEGIN ADDING STYLESHEET TO DOCUMENT -------------------

            // add a new style part to the workbook
            WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();

            // create the stylesheet and assign it to the stylepart.stylesheet property
            stylePart.Stylesheet = CreateStyleSheet();

            // save stylesheet to style part
            stylePart.Stylesheet.Save();

            // --------------------------- END ADDING STYLESHEET TO DOCUMENT ----------------------

            // ----------------------------BEGIN ADDING COLUMNS --------------- -------------------

            // create the columns
            Columns worksheetColumns = CreateWorksheetColumns();

            // append the columns.  NOTE!!! Only works if you provide nothing to the
            // new Worksheet declaration
            worksheetPart.Worksheet.AppendChild(worksheetColumns);

            // save the workbook part
            workbookPart.Workbook.Save();

            // ------------------------------- END ADDING COLUMNS --------------------------------

            // ---------------------------- BEGIN PRELIMINARY BUILD OF WORKSHEET -----------------
            worksheetPart.Worksheet.Append(new SheetData());

            // Add a Sheets collection to the Workbook.
            Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Create a new worksheet associate it with the workbook.
            Sheet sheet = new Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = WorksheetName };

            // add the worksheet to the Sheets collection
            sheets.Append(sheet);

            // save the workbook part
            workbookPart.Workbook.Save();

            // Get the sheetData object.  This is actually the object you will be adding the data
            // rows to.
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // --------------------------- END PRELIMINARY BUILD OF WORKSHEET -------------------

            // ----------------------------- BEGIN BUILDING THE CELLS (ROW BY ROW)---------------

            // first, do the header row
            Row rw = new Row();

            // append the row
            sheetData.Append(rw);

            // now we need to set the row index of the row for
            // cell reference, formula purposes
            rw.RowIndex = UInt32Value.FromUInt32((uint)sheetData.Elements().Count());

            // create the row
            CreateHeaderDataRow(rw);

            // create  a dataset
            List<DataItem> dataItems = CreateDataSet();

            // lets set a row pointer to the beginning row
            int BeginningDataRow = 0;

            // now, create each data row by appending it to sheetdata
            for (int i = 0; i < dataItems.Count; i++)
            {
                // new data row
                rw = new Row();

                // append the row
                sheetData.Append(rw);

                // now we need to set the row index of the row for
                // cell reference and formula purposes
                rw.RowIndex = UInt32Value.FromUInt32((uint)sheetData.ChildElements.Count);

                // create the data row
                CreateDataRow(dataItems[i], rw);

                // track the beginning row because we will use it to buil the total line
                if (i == 0)
                {
                    BeginningDataRow = sheetData.ChildElements.Count;
                }
            }

            // now, we build a total row
            rw = new Row();

            // append to the sheet data
            sheetData.Append(rw);

            // now we need to set the row index of the row for
            // cell reference and formula purposes
            rw.RowIndex = UInt32Value.FromUInt32((uint)sheetData.ChildElements.Count);

            // create the total row.  We know the begin row is 1 because the header is at 0.  We know the
            // we could track the beginning row easily enough but for
            CreateFooterRow(BeginningDataRow, dataItems.Count, rw);

            // ------------------------------- END BUILDING THE DATA CELLS -----------------------

            // ------------------------------- BEGIN CLOSING OUT WORKBOOK ------------------------

            // save the worksheet
            worksheetPart.Worksheet.Save();

            // save the workbook
            workbookPart.Workbook.Save();

            // close the document and flush to stream
            doc.Close();

            // -------------------------------END CLOSING WORKBOOK ------------------------------

            // rewind the memory stream
            ms.Seek(0, SeekOrigin.Begin);

            // return the file stream
            rslt = new FileStreamResult(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            // back to the browser
            return rslt;
        }

        #endregion CompleteWorksheet

        #region CreateWorksheetColumns

        // ******************************************************************************
        // ******************************************************************************
        // CreateWorksheetColumns configures the sizing of columns that
        // comprise the worksheet
        // ******************************************************************************
        // ******************************************************************************
        private Columns CreateWorksheetColumns()
        {
            // define a new columns object
            Columns workSheetColumns = new Columns();

            // invoice number column
            workSheetColumns.Append(CreateColumn(16.0, 1, true));

            // date column
            workSheetColumns.Append(CreateColumn(25.0, 2, true));

            // first name column
            workSheetColumns.Append(CreateColumn(20, 3, true));

            // last name column
            workSheetColumns.Append(CreateColumn(20.0, 4, true));

            // will pickup column
            workSheetColumns.Append(CreateColumn(15.0, 5, true));

            // qty column
            workSheetColumns.Append(CreateColumn(15.0, 6, true));

            // unit price column
            workSheetColumns.Append(CreateColumn(15.0, 7, true));

            // subtotal column
            workSheetColumns.Append(CreateColumn(15.0, 8, true));

            return workSheetColumns;
        }

        private static Column CreateColumn(double width, uint min, bool customWidth)
        {
            Column col = new Column();
            col.Width = DoubleValue.FromDouble(width);
            col.Min = UInt32Value.FromUInt32(min);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(customWidth);
            return col;
        }

        #endregion CreateWorksheetColumns

        #region CreateWorksheetStyleSheet

        // ******************************************************************************
        // ******************************************************************************
        // CreateWorksheetStylesheet creates the stylesheet for the worksheet
        //
        // When creating a stylesheet, the idea is to:
        // First, Define all the number formats, fonts, fills, and borders as separate entities.
        // Then, pick a number format, a font, a fill, and a border and create a cell format.
        // Repeat this for as many cell formats as required in the worksheet. It can be
        // any amount.
        // ******************************************************************************
        // ******************************************************************************
        private Stylesheet CreateStyleSheet()
        {
            // values to easily keep track of entities so when be
            // finally build the cell formats, we do it with something friendly
            uint NUMBERING_FORMAT_DATETIME = 200;
            uint NUMBERING_FORMAT_INVOICE_NUMBER = 201;
            uint NUMBERING_FORMAT_QTY = 202;
            uint NUMBERING_FORMAT_CURRENCY = 203;

            // we are going to set these when we build the fills, that way,
            // we don't have to come back up here everytime and modify the values;
            uint FONTID_DEFAULT = 0;
            uint FONTID_DEFAULT_BOLD = 0;
            uint FONTID_RED_BOLD = 0;

            uint FILLID_DEFAULT = 0;
            uint FILLID_PATTERN_VALUE_GRAY_125 = 0;
            uint FILLID_PATTERN_GOLD = 0;
            uint FILLID_PATTERN_GREEN = 0;

            uint BORDERID_DEFAULT = 0;
            uint BORDERID_GRAY = 1;

            Stylesheet styleSheet = new Stylesheet();

            // define a new numbering format collection.  this collection will hold all the
            // numbering formats used throughout the worksheet(s)

            styleSheet.NumberingFormats = new NumberingFormats();

            // THis is for the date, we give the number format an ID of 200.  Note
            // ID is a UInt32Value
            var numberingFormat = CreateNumberingFormat("mm/dd/yyyy hh:mm:ss", NUMBERING_FORMAT_DATETIME);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this number format is for the invoice number
            numberingFormat = CreateNumberingFormat("00000000", NUMBERING_FORMAT_INVOICE_NUMBER);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this number format is for the quantity
            numberingFormat = CreateNumberingFormat("#", NUMBERING_FORMAT_QTY);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this format is for the unit price and the subtotal columns.
            numberingFormat = CreateNumberingFormat("$#.00", NUMBERING_FORMAT_CURRENCY);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // update the collection count.  Don't know why object library can't do this but it doesn't
            styleSheet.NumberingFormats.Count = UInt32Value.FromUInt32((uint)styleSheet.NumberingFormats.ChildElements.Count);

            // Define a new FOnts collection.  This collection will contain all the fonts
            // used in the worksheet(s).  REMEMBER THE INDEXES on these.  O Based
            styleSheet.Fonts = new Fonts();

            // index 0
            Font font = new Font();         // Default font
            styleSheet.Fonts.Append(font);
            FONTID_DEFAULT = (uint)styleSheet.Fonts.ChildElements.Count - 1;

            // default font bold
            font = CreateFont(bold: true);
            styleSheet.Fonts.Append(font);
            FONTID_DEFAULT_BOLD = (uint)styleSheet.Fonts.ChildElements.Count - 1;

            // index 2. Bold Face Red.  We will use this for the headers
            font = CreateFont(bold: true, hexColor: "FF0000");
            styleSheet.Fonts.Append(font);
            FONTID_RED_BOLD = (uint)styleSheet.Fonts.ChildElements.Count - 1;

            // update the font collection count
            styleSheet.Fonts.Count = UInt32Value.FromUInt32((uint)styleSheet.Fonts.ChildElements.Count);

            // define a fills collection.  Fills are used to create the background and foreground colors.
            // they use another object called the pattern fill.  NOTE!!!! you have to always define the
            // two preset fills.
            styleSheet.Fills = new Fills();

            // Fill Index 0
            styleSheet.Fills.Append(CreateFill(PatternValues.None));
            FILLID_DEFAULT = (uint)styleSheet.Fills.ChildElements.Count - 1;

            // Fill Index 1.  Defaults By Micorosoft
            styleSheet.Fills.Append(CreateFill(PatternValues.Gray125));
            FILLID_PATTERN_VALUE_GRAY_125 = (uint)styleSheet.Fills.ChildElements.Count - 1;

            // Fill Index 2 (Custom - Gold)
            styleSheet.Fills.Append(CreateFill(PatternValues.Solid, hexColor: "f9df02"));
            FILLID_PATTERN_GOLD = (uint)styleSheet.Fills.ChildElements.Count - 1;

            // Fill Index 3 (Custom - Green)
            styleSheet.Fills.Append(CreateFill(PatternValues.Solid, hexColor: "00ff00"));
            FILLID_PATTERN_GREEN = (uint)styleSheet.Fills.ChildElements.Count - 1;

            // update the fills collection count
            styleSheet.Fills.Count = UInt32Value.FromUInt32((uint)styleSheet.Fills.ChildElements.Count);

            // Define the borders used in the worksheets.
            styleSheet.Borders = new Borders();

            // default border
            Border border = new Border();
            styleSheet.Borders.Append(border);
            BORDERID_DEFAULT = (uint)styleSheet.Borders.ChildElements.Count - 1;

            border = CreateBorder(hexColor: "b4b4b4");
            styleSheet.Borders.Append(border);
            BORDERID_GRAY = (uint)styleSheet.Borders.ChildElements.Count - 1;

            // update the borders collection count
            styleSheet.Borders.Count = UInt32Value.FromUInt32((uint)styleSheet.Borders.ChildElements.Count);

            // create a new cell formats collection for the stylesheet
            styleSheet.CellFormats = new CellFormats();

            // index 0 - Default Cell Format
            CellFormat cellFormat = new CellFormat();
            styleSheet.CellFormats.Append(cellFormat);
            CELLSTYLE_DEFAULT = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            //TODO: Continue from Here
            // index 1 (Header For Left Justified Cells)
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top, fontId: FONTID_RED_BOLD, fillId: FILLID_PATTERN_GOLD, borderId: BORDERID_GRAY);
            styleSheet.CellFormats.Append(cellFormat);
            HEADER_CELLSTYLE_LEFT_JUSTIFIED = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // index 2 (Header For Right Justified Cells)
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Right, VerticalAlignmentValues.Top, fontId: FONTID_RED_BOLD, fillId: FILLID_PATTERN_GOLD, borderId: BORDERID_GRAY);
            styleSheet.CellFormats.Append(cellFormat);
            HEADER_CELLSTYLE_RIGHT_JUSTIFIED = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // index 3  TEXT CELLS
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_TEXT = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // (Date)
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top, NUMBERING_FORMAT_DATETIME);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_DATE = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // Invoice Number
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top, NUMBERING_FORMAT_INVOICE_NUMBER);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_INVOICE_NUMBER = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // index 6  WILL PICKUP.  BOOLEAN CELL
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_WILL_PICKUP = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // Qty
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Right, VerticalAlignmentValues.Top, NUMBERING_FORMAT_QTY);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_QTY = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // Unit, Total - Currency
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Right, VerticalAlignmentValues.Top, NUMBERING_FORMAT_CURRENCY);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_CURRENCY = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // Total Label
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Right, VerticalAlignmentValues.Top);
            styleSheet.CellFormats.Append(cellFormat);
            FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // Total All Cells  - Currency
            cellFormat = CreateCellFormat(HorizontalAlignmentValues.Right, VerticalAlignmentValues.Top, NUMBERING_FORMAT_CURRENCY, FONTID_DEFAULT_BOLD, FILLID_PATTERN_GREEN, BORDERID_GRAY);
            styleSheet.CellFormats.Append(cellFormat);
            FOOTER_CELLSTYLE_TOTAL_CURRENCY = (uint)styleSheet.CellFormats.ChildElements.Count - 1;

            // now update th cell formats count
            styleSheet.CellFormats.Count = UInt32Value.FromUInt32((uint)styleSheet.CellFormats.ChildElements.Count);

            return styleSheet;
        }

        private static CellFormat CreateCellFormat(HorizontalAlignmentValues horizontalAli, VerticalAlignmentValues verticalAli, uint? numberingFormat = null, uint? fontId = null, uint? fillId = null, uint? borderId = null)
        {
            var cellFormat = new CellFormat();
            if (numberingFormat != null)
            {
                cellFormat.NumberFormatId = UInt32Value.FromUInt32((uint)numberingFormat);
                cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            }

            if (fontId != null)
            {
                cellFormat.FontId = UInt32Value.FromUInt32((uint)fontId);
                cellFormat.ApplyFont = BooleanValue.FromBoolean(true);
            }

            if (fillId != null)
            {
                cellFormat.FillId = UInt32Value.FromUInt32((uint)fillId);
                cellFormat.ApplyFill = BooleanValue.FromBoolean(true);
            }

            if (borderId != null)
            {
                cellFormat.BorderId = UInt32Value.FromUInt32((uint)borderId);
                cellFormat.ApplyBorder = BooleanValue.FromBoolean(true);
            }

            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            return cellFormat;
        }

        private static Border CreateBorder(string hexColor, BorderStyleValues defaultStyle = BorderStyleValues.Thin, params BorderStyleValues[] borders)
        {
            var defaultBorder = borders.Length > 0 ? borders[0] : defaultStyle;
            var border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = defaultBorder;
            border.LeftBorder.Color = new Color();
            border.LeftBorder.Color.Rgb = HexBinaryValue.FromString(hexColor);
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = borders.Length > 1 ? borders[1] : defaultBorder;
            border.RightBorder.Color = new Color();
            border.RightBorder.Color.Rgb = HexBinaryValue.FromString(hexColor);
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = borders.Length > 2 ? borders[2] : defaultBorder;
            border.BottomBorder.Color = new Color();
            border.BottomBorder.Color.Rgb = HexBinaryValue.FromString(hexColor);
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = borders.Length > 3 ? borders[3] : defaultBorder;
            border.TopBorder.Color = new Color();
            border.TopBorder.Color.Rgb = HexBinaryValue.FromString(hexColor);
            return border;
        }

        private static Fill CreateFill(PatternValues pattern, string? hexColor = null)
        {
            Fill fill = new Fill();
            PatternFill patternFill = new PatternFill();
            patternFill.PatternType = pattern;
            if (!string.IsNullOrEmpty(hexColor))
            {
                patternFill.ForegroundColor = new ForegroundColor();
                patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString(hexColor);
            }
            fill.PatternFill = patternFill;
            return fill;
        }

        private static Font CreateFont(bool bold, string? hexColor = null)
        {
            var font = new Font();
            font.Bold = new Bold();
            font.Bold.Val = BooleanValue.FromBoolean(bold);
            if (!string.IsNullOrEmpty(hexColor))
            {
                font.Color = new Color();
                font.Color.Rgb = HexBinaryValue.FromString(hexColor);
            }
            return font;
        }

        private static NumberingFormat CreateNumberingFormat(string format, uint NUMBERING_FORMAT_DATETIME)
        {
            NumberingFormat numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = StringValue.FromString(format);
            numberingFormat.NumberFormatId = UInt32Value.FromUInt32(NUMBERING_FORMAT_DATETIME);
            return numberingFormat;
        }

        #endregion CreateWorksheetStyleSheet

        #region CreateHeaderDataRow

        // ******************************************************************************
        // ******************************************************************************
        // This gets the header data row.  It is just like any other data row however,
        // this only has column headers values
        // ******************************************************************************
        // ******************************************************************************
        private void CreateHeaderDataRow(Row rw)
        {
            Cell c = populateCell(rw, "INVOICE#", CellValues.String, "A" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            rw.Append(c);

            c = populateCell(rw, "DATE", CellValues.String, "B" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            rw.Append(c);

            c = populateCell(rw, "FIRST", CellValues.String, "C" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            rw.Append(c);

            // last
            c = populateCell(rw, "LAST", CellValues.String, "D" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            rw.Append(c);

            // will pickup
            c = populateCell(rw, "WILL PICKUP", CellValues.String, "E" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            rw.Append(c);

            // qty header
            c = populateCell(rw, "QTY", CellValues.String, "F" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            rw.Append(c);

            c = populateCell(rw, "UNITPRICE", CellValues.String, "G" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            rw.Append(c);

            c = populateCell(rw, "SUBTOTAL", CellValues.String, "H" + rw.RowIndex.ToString(), HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            rw.Append(c);
        }

        private Cell populateCell<T>(Row rw, T value, CellValues dataType, string cellReference, uint styleIndex)
        {
            if (value == null) throw new ArgumentNullException("value");
            var c = new Cell();
            if (dataType == CellValues.Date && value is DateTime)
            {
                c.CellValue = new CellValue((DateTime)(object)value);
            }
            else if (typeof(T).IsPrimitive || typeof(T) == typeof(decimal) || typeof(T) == typeof(string))
            {
                c.CellValue = new CellValue(Convert.ChangeType(value, typeof(T)).ToString());
            }
            else
            {
                throw new InvalidDataException("The type chosen for the Cell is not compatible!");
            }
            c.DataType = dataType;
            c.StyleIndex = UInt32Value.FromUInt32(styleIndex);
            c.CellReference = cellReference;
            return c;
        }

        #endregion CreateHeaderDataRow

        #region CreateDataRow

        // ******************************************************************************
        // ******************************************************************************
        // Create Data Row simply adds the cells to the row for each data record
        // ******************************************************************************
        // ******************************************************************************
        private void CreateDataRow(DataItem itm, Row rw)
        {
            // invoice number (A)
            Cell c = populateCell(rw, itm.InvoiceNumber, CellValues.Number, "A" + rw.RowIndex.ToString(), DATA_CELLSTYLE_INVOICE_NUMBER);
            rw.Append(c);

            // invoice date.  (B)
            c = populateCell(rw, itm.InvoiceDate, CellValues.Date, "B" + rw.RowIndex.ToString(), DATA_CELLSTYLE_DATE);
            rw.Append(c);

            // first name (C)
            c = populateCell(rw, itm.First, CellValues.String, "C" + rw.RowIndex.ToString(), DATA_CELLSTYLE_TEXT);
            rw.Append(c);

            // last name (D)
            c = populateCell(rw, itm.Last, CellValues.String, "D" + rw.RowIndex.ToString(), DATA_CELLSTYLE_TEXT);
            rw.Append(c);

            // will pickup (E).
            // Note, boolean values are "1" (true), "0" (false)"
            // Therefore, if you have a csharp boolean value, you will need to transpose it to
            // one or zero.  We will do that here using the ternary operator.
            //int convertedBooleanValue = itm.WillPickUp ? 1 : 0;
            c = populateCell(rw, itm.WillPickUp, CellValues.Boolean, "E" + rw.RowIndex.ToString(), DATA_CELLSTYLE_WILL_PICKUP);
            rw.Append(c);

            // qty (F)
            c = populateCell(rw, itm.Qty, CellValues.Number, "F" + rw.RowIndex.ToString(), DATA_CELLSTYLE_QTY);
            rw.Append(c);
            Cell QtyCell = c;

            // unit price (G)
            c = populateCell(rw, itm.UnitPrice, CellValues.Number, "G" + rw.RowIndex.ToString(), DATA_CELLSTYLE_CURRENCY);
            rw.Append(c);
            Cell UnitPriceCell = c;

            // build the cell formula object.  The formula is "=F1 * G1"
            // assuming row 1
            CellFormula cellFormula = new CellFormula();
            cellFormula.Text = "=" + QtyCell.CellReference + "*" + UnitPriceCell.CellReference;

            // now build the cell.  NOTE, we don't assign a cell value because
            // that will come from the formula
            c = new Cell();
            c.CellFormula = cellFormula;
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32(DATA_CELLSTYLE_CURRENCY);
            c.CellReference = "H" + rw.RowIndex.ToString();
            rw.Append(c);
        }

        #endregion CreateDataRow

        #region CreateFooterRow

        // ******************************************************************************
        // ******************************************************************************
        // Creates THe Footer Row (total)
        // ******************************************************************************
        // ******************************************************************************
        private void CreateFooterRow(int BeginRow, int DataItemCount, Row rw)
        {
            // invoice number (A)
            Cell c = populateCell(rw, "", CellValues.String, "A" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);

            // invoice date.  (B)
            c = populateCell(rw, "", CellValues.String, "B" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);

            // first name (C)
            c = populateCell(rw, "", CellValues.String, "C" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);

            // last name (D)
            c = populateCell(rw, "", CellValues.String, "D" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);

            c = populateCell(rw, "", CellValues.String, "E" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);

            // qty (F)
            c = populateCell(rw, "", CellValues.String, "F" + rw.RowIndex.ToString(), CELLSTYLE_DEFAULT);
            rw.Append(c);
            Cell QtyCell = c;

            // unit price (G)
            c = populateCell(rw, "Total:", CellValues.String, "G" + rw.RowIndex.ToString(), FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT);
            rw.Append(c);
            Cell UnitPriceCell = c;

            // build the cell formula object.  The formula is SUM(BEGINCELLREF : ENDCELLREF)
            string BeginCellRef = "H" + BeginRow.ToString();
            string EndCellRef = "H" + (BeginRow + DataItemCount - 1).ToString();
            CellFormula cellFormula = new CellFormula();
            cellFormula.Text = "=SUM(" + BeginCellRef + ":" + EndCellRef + ")";

            // now build the cell.  NOTE, we don't assign a cell value because
            // that will come from the formula
            c = new Cell();
            c.CellFormula = cellFormula;
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32(FOOTER_CELLSTYLE_TOTAL_CURRENCY);
            c.CellReference = "H" + rw.RowIndex.ToString();
            rw.Append(c);
        }

        #endregion CreateFooterRow

        #region CreateDataSet

        // ******************************************************************************
        // ******************************************************************************
        // CreateDataSet simply creates a list of data records which is used to build the
        // contents of the excel spreadsheet.
        // ******************************************************************************
        // ******************************************************************************
        private List<DataItem> CreateDataSet()
        {
            List<string> LASTNAMES = new List<string>()
            {
                "Johnson",
                "Earnhardt",
                "Gordon",
                "Petty",
                "Preece",
                "Logano",
                "Keselowski",
                "Trump",
                "Obama",
                "Bush",
                "Clinton",
                "Reagan",
                "Ford",
                "Nixon"
            };

            List<string> FIRSTNAMES = new List<string>()
            {
                "Jimmy",
                "Dale",
                "Jeff",
                "Richard",
                "Ryan",
                "Joey",
                "Brad",
                "Donald",
                "Barak",
                "George",
                "Bill",
                "Ronald",
                "Gerald",
                "Tricky"
            };
            List<DataItem> rt = new List<DataItem>();

            int LASTNAMES_COUNT = LASTNAMES.Count();
            int FIRSTNAMES_COUNT = FIRSTNAMES.Count();
            int ivn = 0;
            DateTime dt = DateTime.Now.AddDays(-365);
            dt = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second);
            Random RND = new Random();

            ivn = RND.Next(1, 5000);

            for (int i = 0; i < 100; i++)
            {
                DataItem itm = new DataItem();
                itm.Last = LASTNAMES[RND.Next(0, LASTNAMES_COUNT)];
                itm.First = FIRSTNAMES[RND.Next(0, FIRSTNAMES_COUNT)];
                itm.InvoiceNumber = ivn + i;

                itm.WillPickUp = (RND.Next(0, 100) > 49) ? true : false;

                itm.Qty = (decimal)RND.Next(1, 100);
                itm.UnitPrice = (decimal)RND.Next(11, 39);
                itm.SubTotal = itm.Qty * itm.UnitPrice;
                itm.InvoiceDate = dt;

                if (i % 5 == 0)
                {
                    dt = dt.Add(new TimeSpan(1, 1, 2, 30));
                }
                else
                {
                    dt = dt.Add(new TimeSpan(2, 3, 35));
                }

                rt.Add(itm);
            }

            return rt;
        }

        #endregion CreateDataSet

        #region DataItem Class

        // ******************************************************************************
        // ******************************************************************************
        // Data Item Class.
        // ******************************************************************************
        // ******************************************************************************

        private class DataItem
        {
            public string First { get; set; } = "";
            public string Last { get; set; } = "";
            public int InvoiceNumber { get; set; } = 0;
            public DateTime InvoiceDate { get; set; } = DateTime.MinValue;
            public Boolean WillPickUp { get; set; } = false;
            public decimal Qty { get; set; } = 0M;
            public decimal UnitPrice { get; set; } = 0M;
            public decimal SubTotal { get; set; } = 0;
        }

        #endregion DataItem Class
    }
}