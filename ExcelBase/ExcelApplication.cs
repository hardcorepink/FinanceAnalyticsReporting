using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.IO;

namespace ExcelBase
{
    public static partial class ExcelApplication
    {

        #region constructor

        static ExcelApplication()
        {
            //static constructor if required
        }

        #endregion constructor
        
        #region fields

        private static WorkbookCollection _workbooks = new WorkbookCollection();

        #endregion fields

        #region properties

        public static bool ScreenUpdating
        {
            get { return (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 40); }
            set { XlCall.Excel(XlCall.xlcEcho, value); }
        }

        public static WorkbookCollection Workbooks
        {
            get { return _workbooks; }
        }

        #endregion properties

        #region nestedClasses

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
                            string filePath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2, wb.Name);
                            string fullFilePath = string.Format($"{filePath}\\{wb.Name}");
                            string pathMatchPotential = Path.GetFullPath(fullFilePath);
                            if (string.Equals(pathLookingFor, pathMatchPotential, StringComparison.OrdinalIgnoreCase)) { return wb; }
                        }
                        catch { }
                    }
                    return null;
                }
            }
            
        }
        #endregion nestedClasses

    }



}









