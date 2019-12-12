using System;
using Microsoft.Office.Interop.Excel;

namespace TTTOfficeWrapper
{
    public class ExcelHelper
    {
        Application app;
        Workbook wb;
        Worksheet ws;
        Range rg;
        Interior rint;
        Borders rbdrs;
        Border rbdr;
        Range rgrow;
        Chart chrt;
        PivotCaches pivcaches;
        PivotCache pivcache;
        PivotTable pivtbl;
        PivotFields pivflds;
        PivotField pivfld;
        int[][] fieldinfo;

        public void Application_Instantiate()
        {
            app = new Application();
        }

        public string Application_International(int idx)
        {
            return (string)app.International[idx];
        }

        public string Application_Version()
        {
            return app.Version;
        }

        public void FieldinfoArray_Create(int noOfCols)
        {
            fieldinfo = new int[noOfCols][];
        }

        public void FieldinfoArray_Set(int colNo, int fieldType)
        {
            fieldinfo[colNo - 1] = new int[] { colNo, fieldType };
        }

        public void Application_OpenText(string filename)
        {
            app.Workbooks.OpenText(
                Filename: filename,
                Origin: 2, // 1 = Macintosh, 2 = Microsoft Windows, 3 = MS-DOS
                DataType: XlTextParsingType.xlDelimited,
                TextQualifier: XlTextQualifier.xlTextQualifierNone,
                ConsecutiveDelimiter: false,
                Tab: true,
                Semicolon: false,
                Comma: false,
                Space: false,
                Other: false,
                FieldInfo: fieldinfo,
                TrailingMinusNumbers: true
                );
        }

        public void Application_WorkbookAdd()
        {
            wb = app.Workbooks.Add();
            ws = (Worksheet)wb.Worksheets.Item[1];
        }

        public void Application_ActiveWorkbook()
        {
            wb = app.ActiveWorkbook;
            ws = (Worksheet)wb.Worksheets.Item[1];
        }

        public void Application_SetSheet()
        {
            ws = (Worksheet)wb.Worksheets.Item[1];
        }

        public void Application_AddInfo(string sheetName, string bookTitle, string bookAuthor)
        {
            ws.Name = sheetName;
            wb.Title = bookTitle;
            wb.Author = bookAuthor;
        }

        public void Application_Visible(bool visible)
        {
            app.Visible = visible;
        }

        public void Application_UserControl(bool userControl)
        {
            app.UserControl = userControl;
        }

        public void Application_UseSystemSeparators()
        {
            app.UseSystemSeparators = true;
        }

        public void Worksheet_Activate()
        {
            ws.Activate();
        }

        public void Worksheet_SetRange(string fromCell, string toCell)
        {
            rg = ws.Range[fromCell, toCell];
        }

        public void Range_SetFontBold(bool bold)
        {
            rg.Font.Bold = bold;
        }

        public void Range_SetFontItalic(bool italic)
        {
            rg.Font.Italic = italic;
        }

        public void Range_SetFontSize(int size)
        {
            rg.Font.Size = size;
        }

        public void Range_SetFontColor(int color)
        {
            rg.Font.Color = color;
        }

        public void Range_SetBackgroundColor(int color)
        {
            rint = rg.Interior;
            rint.Color = color;
        }

        public void Range_SetInteriorColorIndex(int idx)
        {
            rint = rg.Interior;
            rint.ColorIndex = idx;
        }

        public void Range_SetBorderTop()
        {
            rbdrs = rg.Borders;
            rbdr = rbdrs.Item[XlBordersIndex.xlEdgeTop];
            rbdr.LineStyle = XlLineStyle.xlContinuous;
            rbdr.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            rbdr.TintAndShade = 0;
            rbdr.Weight = 2;
        }

        public void Range_SetBorderBottom()
        {
            rbdrs = rg.Borders;
            rbdr = rbdrs.Item[XlBordersIndex.xlEdgeBottom];
            rbdr.LineStyle = XlLineStyle.xlContinuous;
            rbdr.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            rbdr.TintAndShade = 0;
            rbdr.Weight = 2;
        }

        public void Range_SetBorderLeft()
        {
            rbdrs = rg.Borders;
            rbdr = rbdrs.Item[XlBordersIndex.xlEdgeLeft];
            rbdr.LineStyle = XlLineStyle.xlContinuous;
            rbdr.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            rbdr.TintAndShade = 0;
            rbdr.Weight = 2;
        }

        public void Range_SetBorderRight()
        {
            rbdrs = rg.Borders;
            rbdr = rbdrs.Item[XlBordersIndex.xlEdgeRight];
            rbdr.LineStyle = XlLineStyle.xlContinuous;
            rbdr.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            rbdr.TintAndShade = 0;
            rbdr.Weight = 2;
        }

        public void Range_SetWrapText()
        {
            rg.WrapText = true;
        }

        public void Range_SetColumnWidth(int width)
        {
            rg.EntireColumn.ColumnWidth = width;
        }

        public void Range_SetNumberFormat(string fmt)
        {
            rg.NumberFormat = fmt;
        }

        public void Range_SetValue(string val)
        {
            rg.Value2 = val;
        }

        public string Range_GetValue()
        {
            return rg.Value2.ToString();
        }

        public void Range_EntireRowMergeCells()
        {
            rg.EntireRow.MergeCells = true;
        }

        public void Sheet_ColumnsAutoFit()
        {
            ws.Columns.AutoFit();
        }

        public void Range_DeleteEntireRow()
        {
            rgrow = rg.EntireRow;
            rgrow.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        public void Workbook_ChartsAdd()
        {
            chrt = (Chart)wb.Charts.Add();
        }

        public void Chart_SetSourceData()
        {
            chrt.SetSourceData(
                Source: rg
                );
        }

        public void Chart_SetType(int chartType)
        {
            chrt.ChartType = (XlChartType)chartType;
        }

        public void Chart_SetChartTitleAndLocation(string title)
        {
            chrt.ChartTitle.Text = title;
            chrt.Location(XlChartLocation.xlLocationAsObject, ws.Name);
        }

        public void Workbook_CreatePivot(string firstCell, string lastCell)
        {
            pivcaches = wb.PivotCaches();
            pivcache = pivcaches.Create(
                SourceType: XlPivotTableSourceType.xlDatabase,
                SourceData: firstCell + ":" + lastCell
                );
            pivtbl = pivcache.CreatePivotTable(
                TableDestination: "",
                TableName: "SchurTable"
                );
            pivflds = (PivotFields)pivtbl.PivotFields();
        }

        public void Pivot_SetField()
        {
            pivfld = (PivotField)pivflds.Item(rg.Value2);
        }

        public void PivotTable_AddDataField(string caption, int func, string decFormat)
        {
            pivtbl.AddDataField(pivfld, caption, func).NumberFormat = decFormat;
        }

        public void PivotField_SetOrientationRow()
        {
            pivfld.Orientation = XlPivotFieldOrientation.xlRowField;
        }

        public void PivotField_SetOrientationColumn()
        {
            pivfld.Orientation = XlPivotFieldOrientation.xlColumnField;
        }

        public void PivotField_SetPosition()
        {
            pivfld.Position = 1;
        }

        public void Workbook_SaveAs(string filename)
        {
            wb.SaveAs(
                Filename: filename,
                FileFormat: 51
                );
        }

        public void Workbook_Close()
        {
            wb.Close(
                SaveChanges: true
                );
        }

        public void Workbook_CloseNoSave()
        {
            wb.Close(
                SaveChanges: false
                );
        }

        public void Application_Quit()
        {
            app.Quit();
        }
    }
}

