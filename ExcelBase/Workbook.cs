using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.ComponentModel;

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
        /// Creates a new workbook and returns constructed workbook.
        /// Throws an error and does not construct if cannot create.
        /// </summary>
        public Workbook(Boolean newWorkbook)
        {
            try
            {
                if (newWorkbook)
                {
                    XlCall.Excel(XlCall.xlcNew, 5);
                }

                this._fileName = (string)XlCall.Excel(XlCall.xlfGetDocument, 88);
            }

            catch { throw; }
            //this is a new workbook so we do not have a file path 
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
        private static string AttemptOpen(string fullFilePath, bool checkIfAlreadyOpen = true,
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
                string workbookOwningWindowName = (string)XlCall.Excel(XlCall.xlfGetWindow, 1, window_text);
                if (workbookOwningWindowName == this.Name)
                {
                    XlCall.Excel(XlCall.xlcActivate, window_text, Type.Missing);
                    XlCall.Excel(XlCall.xlcHide);
                }
            }

            //lets first try just hiding the window by making it's size 0

        }

        public static string WorkbookNameFromSquareBrackets(string squareBracketFulName)
        {
            
            int index = squareBracketFulName.IndexOf("]");
            if (index > 0)
                squareBracketFulName = squareBracketFulName.Substring(0, index);

            return squareBracketFulName.Replace("[", "");
        }

        #region properties
        public string Name { get => _fileName; }




        #endregion properties

    }
}
