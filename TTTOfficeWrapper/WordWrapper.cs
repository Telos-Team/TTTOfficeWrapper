using System;
using Microsoft.Office.Interop.Word;

namespace TTTOfficeWrapper
{
    public static class WordWrapper
    {
        #region CallOpen
        public static Document CallOpen(ref Application application, string fileName)
        {
            return application.Documents.Open(
                FileName: fileName
                );
        }
        public static Document CallOpen(ref Application application, string fileName, bool confirmConversions, bool readOnly)
        {
            return application.Documents.Open(
                FileName: fileName, 
                ConfirmConversions: confirmConversions, 
                ReadOnly: readOnly
                );
        }
        #endregion // CallOpen

        #region CallOpen2
        public static Document CallOpen2(ref Application application, string fileName, ref string protectiontype)
        {
            Document document = application.Documents.Open(
                FileName: fileName
                );
            protectiontype = "";
            switch(document.ProtectionType)
            {
                case WdProtectionType.wdAllowOnlyComments:
                    protectiontype = "Allow Only Comments";
                    break;
                case WdProtectionType.wdAllowOnlyFormFields:
                    protectiontype = "Allow Only Form Fields";
                    break;
                case WdProtectionType.wdAllowOnlyReading:
                    protectiontype = "Allow Only Reading";
                    break;
                case WdProtectionType.wdAllowOnlyRevisions:
                    protectiontype = "Allow Only Revisions";
                    break;
                case WdProtectionType.wdNoProtection:
                    protectiontype = "No Protection";
                    break;
            }
            return document;
        }
        public static Document CallOpen2(ref Application application, string fileName, bool confirmConversions, bool readOnly, ref string protectiontype)
        {
            Document document = application.Documents.Open(
                FileName: fileName,
                ConfirmConversions: confirmConversions,
                ReadOnly: readOnly
                );
            protectiontype = "";
            switch (document.ProtectionType)
            {
                case WdProtectionType.wdAllowOnlyComments:
                    protectiontype = "Allow Only Comments";
                    break;
                case WdProtectionType.wdAllowOnlyFormFields:
                    protectiontype = "Allow Only Form Fields";
                    break;
                case WdProtectionType.wdAllowOnlyReading:
                    protectiontype = "Allow Only Reading";
                    break;
                case WdProtectionType.wdAllowOnlyRevisions:
                    protectiontype = "Allow Only Revisions";
                    break;
                case WdProtectionType.wdNoProtection:
                    protectiontype = "No Protection";
                    break;
            }
            return document;
        }
        #endregion // CallOpen2

        #region CallOpen3
        public static Document CallOpen3(ref Application application, string fileName, ref string protectiontype, string pwd)
        {
            Document document = application.Documents.Open(
                FileName: fileName
                );
            protectiontype = "";
            switch (document.ProtectionType)
            {
                case WdProtectionType.wdAllowOnlyComments:
                    protectiontype = "Allow Only Comments";
                    break;
                case WdProtectionType.wdAllowOnlyFormFields:
                    protectiontype = "Allow Only Form Fields";
                    document.Unprotect(pwd);
                    break;
                case WdProtectionType.wdAllowOnlyReading:
                    protectiontype = "Allow Only Reading";
                    break;
                case WdProtectionType.wdAllowOnlyRevisions:
                    protectiontype = "Allow Only Revisions";
                    break;
                case WdProtectionType.wdNoProtection:
                    protectiontype = "No Protection";
                    break;
            }
            return document;
        }
        public static Document CallOpen3(ref Application application, string fileName, bool confirmConversions, bool readOnly, ref string protectiontype, string pwd)
        {
            Document document = application.Documents.Open(
                FileName: fileName,
                ConfirmConversions: confirmConversions,
                ReadOnly: readOnly
                );
            protectiontype = "";
            switch (document.ProtectionType)
            {
                case WdProtectionType.wdAllowOnlyComments:
                    protectiontype = "Allow Only Comments";
                    break;
                case WdProtectionType.wdAllowOnlyFormFields:
                    protectiontype = "Allow Only Form Fields";
                    document.Unprotect(pwd);
                    break;
                case WdProtectionType.wdAllowOnlyReading:
                    protectiontype = "Allow Only Reading";
                    break;
                case WdProtectionType.wdAllowOnlyRevisions:
                    protectiontype = "Allow Only Revisions";
                    break;
                case WdProtectionType.wdNoProtection:
                    protectiontype = "No Protection";
                    break;
            }
            return document;
        }
        #endregion // CallOpen3

        #region CallQuit
        public static void CallQuit(ref Application application)
        {
            application.Quit();
        }
        public static void CallQuit(ref Application application, bool saveChanges)
        {
            application.Quit(
                SaveChanges: saveChanges
                );
        }
        #endregion // CallQuit

        #region CallAdd
        public static Document CallAdd(ref Application application)
        {
            return application.Documents.Add();
        }
        #endregion CallAdd

        #region CallRowsAdd
        public static void CallRowsAdd(ref Rows rows)
        {
            rows.Add();
        }
        #endregion CallRowsAdd

        #region CallTableCellRangeDelete
        public static void CallTableCellRangeDelete(ref Document document, int tableNo, int rowNo, int columnNo)
        {
            document.Tables[tableNo].Cell(rowNo, columnNo).Range.Delete();
        }
        #endregion CallTableCellRangeDelete

        #region CallRangeDelete
        public static void CallRangeDelete(ref Range range)
        {
            range.Delete();
        }
        public static string CallRangeDelete2(ref Object range)
        {
            return range.GetType().ToString();
            //((Range)range).Delete();
        }
        #endregion CallRangeDelete

        #region CallAddPicture
        public static InlineShape CallAddPicture(ref InlineShapes inlineShapes, string fileName, bool linkToFile, bool saveWithDocument)
        {
            return inlineShapes.AddPicture(
                FileName: fileName,
                LinkToFile: linkToFile,
                SaveWithDocument: saveWithDocument
                );
        }
        #endregion CallAddPicture

        #region CallRange
        public static Range CallRange(ref Document document)
        {
            return document.Range();
        }
        #endregion CallRange

        #region CallRangeGetSpellingSuggestions
        public static SpellingSuggestions CallRangeGetSpellingSuggestions(ref Range range)
        {
            return range.GetSpellingSuggestions();
        }
        #endregion CallRangeGetSpellingSuggestions

        // #region CallSaveAs
        // public static void CallSaveAs(ref Document document, string fileName)
        // {
        //     document.SaveAs2(
        //         FileName: fileName
        //         );
        // }
        // #endregion CallSaveAS

        #region CallClose
        public static void CallClose(ref Document document, bool saveChanges)
        {
            document.Close(
                SaveChanges: saveChanges
                );
        }
        #endregion CallClose

        // #region GetProtectionTypeAsText
        // public static string GetProtectionTypeAsText(ref Document document)
        // {
        //    if (document != null)
        //    {
        //        if(document.ProtectionType.Equals(WdProtectionType.wdAllowOnlyFormFields))
        //        {
        //            return "Allow Only Form Fields";
        //        }
        //        /*
        //        switch(document.ProtectionType)
        //        {
        //            case WdProtectionType.wdAllowOnlyComments:
        //                return "Allow Only Comments";
        //            case WdProtectionType.wdAllowOnlyFormFields:
        //                return "Allow Only Form Fields";
        //            case WdProtectionType.wdAllowOnlyReading:
        //                return "Allow Only Reading";
        //            case WdProtectionType.wdAllowOnlyRevisions:
        //                return "Allow Only Revisions";
        //            case WdProtectionType.wdNoProtection:
        //                return "No Protection";
        //        }
        //        */
        //        return "Unkonwn Protection Type";
        //    }
        //    else
        //    {
        //        return "Unknown Document";
        //    }
        // }
        // #endregion GetProtectionTypeAsText

        #region CallProtect
        public static void CallProtect(ref Document document, WdProtectionType protectionType)
        {
            document.Protect(
                Type: protectionType
                );
        }
        #endregion CallProtect

        #region CallProtect2
        public static void CallProtect2(ref Document document, bool noReset, string pwd)
        {
            document.Protect(
                Type: WdProtectionType.wdAllowOnlyFormFields,
                NoReset: noReset,
                Password: pwd
                );
        }
        #endregion CallProtect2

        #region CallUnprotect
        public static void CallUnprotect(ref Document document)
        {
            document.Unprotect();
        }
        #endregion CallUnprotect

        #region CallUnprotect2
        public static void CallUnprotect2(ref Document document, string pwd)
        {
            document.Unprotect(
                Password: pwd
                );
        }
        #endregion CallUnprotect2

        #region CallPrintOut
        public static void CallPrintOut(ref Document document, bool background)
        {
            document.PrintOut(
                Background: background
                );
        }
        #endregion CallPrintOut
    }
}
