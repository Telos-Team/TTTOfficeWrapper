using System;
using Microsoft.Office.Interop.Excel;

namespace TTTOfficeWrapper
{
    public static class ExcelWrapper
    {
        #region CallOpen
        public static Workbook CallOpen(ref Application application, string fileName)
        {
            return application.Workbooks.Open(
                Filename: fileName
                );
        }
        public static Workbook CallOpen(ref Application application, string fileName, bool readOnly)
        {
            return application.Workbooks.Open(
                Filename: fileName,
                ReadOnly: readOnly
                );
        }
        #endregion // CallOpen

        #region CallQuit
        public static void CallQuit(ref Application application)
        {
            application.Quit();
        }
        #endregion CallQuit

        #region CallRun
        public static void CallRun(ref Application application, string macro)
        {
            application.Run(
                Macro: macro
                );
        }
        #endregion CallRun

        #region MakeVBEPart
        public static Application MakeVBEPart(Application application, string macroFilename, string macroName)
        {
            Workbook wb = application.Workbooks.Add();
            Microsoft.Vbe.Interop.VBE vbe = application.VBE;
            Microsoft.Vbe.Interop.VBProject proj = vbe.ActiveVBProject;
            Microsoft.Vbe.Interop.VBComponent comp = proj.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
            comp.Name = "TTT";
            Microsoft.Vbe.Interop.CodeModule codemod = comp.CodeModule;
            codemod.AddFromFile(macroFilename);
            application.Run(macroName);
            proj.VBComponents.Remove(comp);
            wb.Close(false);
            return application;
        }
        #endregion MakeVBEPart

        #region CallAdd
        public static Workbook CallAdd(ref Application application)
        {
            return application.Workbooks.Add();
        }
        #endregion CallAdd

        #region CallChartsAdd
        public static Chart CallChartsAdd(ref Workbook workbook)
        {
            return (Chart) workbook.Charts.Add();
        }
        public static Chart CallChartsAdd2(ref Application application)
        {
            return (Chart) application.Charts.Add();
        }
        #endregion CallChartsAdd

        #region CallChartSetSourceData
        public static void CallChartSetSourceData(ref Chart chart, Range range)
        {
            chart.SetSourceData(
                Source: range
                );
        }
        #endregion CallChartSetSourceData

        #region CallPivotCachesCreate
        public static PivotCache CallPivotCachesCreate(ref PivotCaches pivotCaches, XlPivotTableSourceType sourceType, string sourceData)
        {
            return pivotCaches.Create(
                SourceType: sourceType,
                SourceData: sourceData
                );
        }
        #endregion CallPivotcachesCreate

        #region CallCreatePivotTable
        public static PivotTable CallCreatePivotTable(ref PivotCache pivotCache, string tableDestination, string tableName)
        {
            return pivotCache.CreatePivotTable(
                TableDestination: tableDestination,
                TableName: tableName
                );
        }
        #endregion CallCreatePivotTable

        #region CallRangeEntireRowDelete
        public static void CallRangeEntireRowDelete(ref Range range)
        {
            range.EntireRow.Delete();
        }
        #endregion CallRangeEntireRowDelete

        #region CallRangeDelete
        public static void CallRangeDelete(ref Range range)
        {
            range.Delete();
        }
        #endregion CallRangeDelete

        #region CallPivotFields
        public static PivotFields CallPivotFields(ref PivotTable pivotTable)
        {
            return (PivotFields) pivotTable.PivotFields();
        }
        #endregion CallPivotFields

        #region CallSaveAs
        // public static void CallSaveAs(ref Workbook workbook, string fileName)
        // {
        //     workbook.SaveAs(
        //         Filename: fileName
        //         );
        // }
        public static void CallSaveAs2(ref Workbook workbook, string fileName, XlFileFormat fileFormat)
        {
            workbook.SaveAs(
                Filename: fileName,
                FileFormat: fileFormat
                );
        }
        #endregion CallSaveAs

        #region CallClose
        public static void CallClose(ref Workbook workbook)
        {
            workbook.Close();
        }
        public static void CallClose(ref Workbook workbook, bool saveChanges)
        {
            workbook.Close(
                SaveChanges: saveChanges
                );
        }
        #endregion CallClose

        #region CallPrintOut
        public static void CallPrintOut(ref Workbook workbook)
        {
            workbook.PrintOutEx();
        }
        #endregion CallPrintOut

        #region DoFunctions
        public static void DoGetActiveWorkbook(ref Application application, ref Workbook workbook)
        {
            workbook = application.ActiveWorkbook;
        }

        public static void DoGetSheet(ref Workbook workbook, int itemNo, ref Worksheet worksheet)
        {
            worksheet = (Worksheet) workbook.Worksheets.Item[itemNo];
        }

        public static void DoGetRange(ref Worksheet worksheet, string fromCell, string toCell, ref Range range)
        {
            range = worksheet.Range[fromCell, toCell];
        }

        public static void DoSetDemo(ref Application application)
        {
            _Workbook workbook = application.ActiveWorkbook;
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item[1];
            Range range = worksheet.Range["A4", "D6"];
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Color = 0;
        }

        public static void DoSetFontBold(ref Range range, bool setBold)
        {
            range.Font.Bold = setBold;
        }

        public static void DoSetFontItalic(ref Range range, bool setItalic)
        {
            range.Font.Italic = setItalic;
        }

        public static void DoSetFontSize(ref Range range, int setSize)
        {
            range.Font.Size = setSize;
        }

        public static void DoSetFontColor(ref Range range, int setColor)
        {
            range.Font.Color = setColor;
        }
        #endregion DoFunctions
    }
}
