using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class Workbook
    {
        #region fields        
        private string _workbookName;
        #endregion fields

        #region properties
        public string Name { get => _workbookName; }

        public FileInfo FileInfo
        {
            get
            {
                try
                {
                    string filePath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2, this.Name);
                    string fullFilePath = string.Format($"{filePath}\\{Name}");
                    return new FileInfo(fullFilePath);
                }
                catch { return null; }
            }
        }


        #endregion properties

        #region constructors

        /// <summary>
        /// Creates a workbook instance given the name of the workbook
        /// </summary>       
        public Workbook(string WorkbookName)
        {
            this._workbookName = WorkbookName;
        }

        /// <summary>
        /// Creates a workbook instance and defaults to the current active workbook.
        /// </summary>
        public Workbook()
        {
            try
            {
                this._workbookName = (string)XlCall.Excel(XlCall.xlfGetWorkbook, 16);
            }
            catch { throw; }
        }

        #endregion constructors       


        public void HideAllWorkbookWindows()
        {
            //get an array of all windows, remember first entry into array is always the active window
            object[,] arrayOfWindows = (object[,])XlCall.Excel(XlCall.xlfWindows, 3, Type.Missing);

            //loop through arrayOfWindows and check which workbook each belongs to, if belongs to this workbook
            //then activate the window and hide
            int numberOfWindows = arrayOfWindows.GetLength(1);

            for (int i = 0; i < numberOfWindows; i++)
            {
                string window_text = (string)arrayOfWindows[0, i];

                string workbookOwningWindowName = WorkbookNameFromSquareBrackets
                    ((string)XlCall.Excel(XlCall.xlfGetWindow, 1, window_text));

                if (workbookOwningWindowName == this.Name)
                {
                    XlCall.Excel(XlCall.xlcActivate, window_text, Type.Missing);
                    XlCall.Excel(XlCall.xlcHide);
                }
            }

        }

        public static string WorkbookNameFromSquareBrackets(string squareBracketFulName)
        {

            int index = squareBracketFulName.IndexOf("]");
            if (index > 0)
                squareBracketFulName = squareBracketFulName.Substring(0, index);

            return squareBracketFulName.Replace("[", "");
        }



        public WorksheetsCollection Worksheets { get => new WorksheetsCollection(this.Name); }

        public class WorksheetsCollection : IEnumerable<Worksheet>
        {
            private string _workbookName;

            public WorksheetsCollection(string WorkbookName)
            {
                this._workbookName = WorkbookName;
            }

            public Worksheet this[string worksheetName]
            {
                get
                {
                    try
                    {   //Get Workbook given 1 returns array of worksheet names in format [BookName]SheetName                     
                        object[,] ArrayFullWorksheetNames = (object[,])(XlCall.Excel(XlCall.xlfGetWorkbook, 1, this._workbookName));
                        long numberSheets = ArrayFullWorksheetNames.GetLongLength(1);
                        for (int i = 0; i < numberSheets; i++)
                        {
                            //get just the sheet name
                            string sheetName = Worksheet.WorksheetNameFromFullReference((string)ArrayFullWorksheetNames[0, i]);
                            //compare to what we provided
                            if (String.Equals(worksheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    //construct a new worksheet type from the full worksheet name
                                    return new Worksheet((string)ArrayFullWorksheetNames[0, i]);
                                }
                                catch
                                {
                                    return null;
                                }
                            }
                        }

                        return null;

                    }
                    catch { return null; }
                    //this will loop through workbook worksheets. 

                }
            }

            public IEnumerator<Worksheet> GetEnumerator()
            {
                object[,] ArrayFullWorksheetNames = (object[,])(XlCall.Excel(XlCall.xlfGetWorkbook, 1, this._workbookName));
                long numberSheets = ArrayFullWorksheetNames.GetLongLength(1);
                for (int i = 0; i < numberSheets; i++)
                {
                    //always need to try constructing a worksheet, just in case it can't work (chart sheet for excample)
                    Worksheet WorksheetToReturn;
                    try
                    {
                        WorksheetToReturn = new Worksheet((string)ArrayFullWorksheetNames[0, i]);
                    }
                    catch
                    {
                        continue;
                    }
                    yield return WorksheetToReturn;

                }
            }
            IEnumerator IEnumerable.GetEnumerator()
            {
                return this.GetEnumerator();
            }
        }
    }


    public class WorkbookCollection : IEnumerable<Workbook>
    {

        public IEnumerator<Workbook> GetEnumerator()
        {
            //get the collection of workbooks
            object[,] workbooksArray = (object[,])XlCall.Excel(XlCall.xlfDocuments);
            int numWorkbooks = workbooksArray.GetLength(1);
            for (int i = 0; i < numWorkbooks; i++)
            {
                yield return new Workbook((string)workbooksArray[0, i]);
            }

            yield break;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }


        /// <summary>
        /// Creates a new workbook, and returns handle of the new workbook.
        /// </summary>
        /// <returns></returns>
        public Workbook Add()
        {
            XlCall.Excel(XlCall.xlcNew, 5);
            string newWorkbookName = (string)XlCall.Excel(XlCall.xlfGetDocument, 88);
            return new Workbook(newWorkbookName);
        }

        /// <summary>
        /// Opens a new workbook given a string of filePath. If the workbook is already open returns an instance of that workbook.
        /// Will return handle to workbook if successful, otherwise throws. Returns null if all else fails.
        /// </summary>            
        public Workbook Open(string fullFilePath, ExcelEnums.XlUpdateLinks update_links = ExcelEnums.XlUpdateLinks.Never,
            bool read_only = false, string password = null)
        {
            //no need to see if workbook already open as when calling xlcOpen automatically makes the workbook active if already open.
            try
            {
                object openFlag = XlCall.Excel(XlCall.xlcOpen, fullFilePath, (int)update_links, read_only, Type.Missing, password, Type.Missing, true, 2,
                                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                if (openFlag is bool) { if ((bool)openFlag) { return new Workbook(); } }
                return null;
            }
            catch { throw; }

        }

        /// <summary>
        /// Returns a workbook instance given the name of a workbook. If the name of the workbook is not found returns null.
        /// There may be multiple workbooks of the same name, if looking for a unique use the FileInfo Indexer of Workbooks.
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Workbook this[string key]
        {
            get
            {
                foreach (Workbook wb in this)
                {
                    if (string.Equals(wb.Name, key, StringComparison.OrdinalIgnoreCase)) return wb;
                }
                return null;
            }
        }

        /// <summary>
        /// Returns a Workbook instance if finds a workbook with matching file path. Else returns a null Workook.
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Workbook this[FileInfo key]
        {
            get
            {
                string pathLookingFor = Path.GetFullPath(key.FullName);
                foreach (Workbook wb in this)
                {
                    //get the full file name of the wb
                    try
                    {
                        if (wb.FileInfo != null)
                        {
                            if (string.Equals(pathLookingFor, wb.FileInfo.FullName, StringComparison.OrdinalIgnoreCase)) { return wb; }
                        }
                    }
                    catch { }
                }
                return null;
            }
        }

    }
}
