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
        private string _fullPath;
        private string _fileName;

        #endregion fields

        #region constructors
        /// <summary>
        /// Given the full file path of the workbook, will check if open already. If not attempts to open.
        /// Throws an error and does not construct if cannot open.
        /// </summary>
        /// <param name="fullFilePath"></param>
        public Workbook(string fullFilePath)
        {
            try
            {
                this._fullPath = Workbook.AttemptOpen(fullFilePath);
                this._fileName = Path.GetFileName(fullFilePath);
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Creates a new workbook if passed true and returns constructed workbook.
        /// If passed false will return current active workbook. Throws an error and does not construct if cannot create.
        /// </summary>
        public Workbook(Boolean newWorkbook)
        {
            try
            {
                if (newWorkbook) { XlCall.Excel(XlCall.xlcNew, 5); }
                this._fileName = (string)XlCall.Excel(XlCall.xlfGetDocument, 88);
            }

            catch { throw; }
            //this is a new workbook so we do not have a file path 
        }

        /// <summary>
        /// Only use this constructor if have workbook name, and are sure that the workbook is currently open in Excel
        /// </summary>
        /// <param name="WorkbookName"></param>
        /// <param name="confirmAlreadyOpened"></param>
        public Workbook(string WorkbookName, Boolean confirmAlreadyOpened)
        {
            try
            {
                List<string> listWorkbooks = ExcelBase.Application.ListWorkbookNames();
                if (listWorkbooks.Exists(x => x == WorkbookName))
                {
                    //if there is an error here there is no path Get Document returns Error if the workbook is not yet saved
                    object fullPathTry = XlCall.Excel(XlCall.xlfGetDocument, 2, WorkbookName);
                    if (fullPathTry is ExcelDna.Integration.ExcelError)
                    {
                        this._fileName = WorkbookName;
                    }
                    else
                    {
                        this._fileName = WorkbookName;
                        this._fullPath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2, WorkbookName);
                    }
                }
                else
                {
                    throw new Exception("Could not find workbook name as passed to workbook constructor");
                }
            }

            catch { throw; }
        }
        #endregion constructors

        /// <summary>
        /// Given a Full File Path will looop through open workbooks in the excel application
        /// and try and find a match. If finds a match to the fullFilePath returns true, otherwise returns false.
        /// </summary>  
        private static Boolean IsWorkbookOpen(string fullFilePath)
        {
            //first get just the filename from fullFilePath
            string fileNameWithExtension = Path.GetFileName(fullFilePath);

            //get an array of all workbooks open. 3 gives the names of all open workbooks (including add-ons)
            //type.missing indicates there is no matching string
            object openWorkbooks = XlCall.Excel(XlCall.xlfDocuments, 3, Type.Missing);

            //explicitly cast return to an array
            object[,] openWorkbooksArray = (object[,])openWorkbooks;
            int numberOfWorkbooks = openWorkbooksArray.GetLength(1);

            //loop through array and find candidates for a file path match
            for (int i = 0; i < numberOfWorkbooks; i++)
            {

                //compare to our fileNameWithoutExtension
                if (fileNameWithExtension == (string)openWorkbooksArray[0, i])
                {
                    //the name of the workbooks matches the name of the file -> now get the file address
                    //and compare
                    string workbookPath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2, fileNameWithExtension);
                    workbookPath = workbookPath + @"\" + fileNameWithExtension;
                    if (workbookPath == fullFilePath) return true;
                }
                System.Diagnostics.Debug.WriteLine(openWorkbooksArray[0, i]);
            }

            return false;

        }

        /// <summary>
        /// Given full file path will attempt to open a workbook. Will rethrow if there is an exception,
        /// returns a Workbook object if open is successful.
        /// </summary>
                 static string AttemptOpen(string fullFilePath, bool checkIfAlreadyOpen = true,
            ExcelEnums.XlUpdateLinks update_links = ExcelEnums.XlUpdateLinks.Never,
            bool read_only = false, string password = null)
        {
            if (checkIfAlreadyOpen)
            {
                if (Workbook.IsWorkbookOpen(fullFilePath))
                {
                    return fullFilePath;
                }
            }

            //could not find open workbook in excel application
            try
            {
                XlCall.Excel(XlCall.xlcOpen, fullFilePath, (int)update_links, read_only, Type.Missing, password, Type.Missing, true, 2,
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                return fullFilePath;
            }
            catch
            {
                throw;
            }

        }

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



        #region properties
        public string Name { get => _fileName; }




        #endregion properties

    }
}
