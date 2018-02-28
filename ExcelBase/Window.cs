using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class Window
    {
        private string _windowName;

        public Window(string windowName)
        {
            this._windowName = windowName;
        }

        public void Hide()
        {
            this.Activate();
            XlCall.Excel(XlCall.xlcHide);
        }

        public void Unhide()
        {
            this.Activate();
            XlCall.Excel(XlCall.xlcUnhide);
        }

        public void Activate()
        {
            XlCall.Excel(XlCall.xlcActivate, this._windowName);
        }

        #region properties


        public int Number
        {
            get { return (int)XlCall.Excel(XlCall.xlfGetWindow, 2, this._windowName); }

        }

        public string WorkbookName
        {
            get
            {
                string workbookText = (string)XlCall.Excel(XlCall.xlfGetWindow, 1, this._windowName);
                //now we have the workbook name in format [Book1]Sheet1       

                if (workbookText.Contains("[") && workbookText.Contains("]"))
                {
                    int startString = workbookText.IndexOf('[');
                    int endString = workbookText.IndexOf(']');
                    return workbookText.Substring(startString + 1, endString - startString - 1);
                }
                else { return workbookText; }

            }
        }
    }





    #endregion properties




    public class WindowsCollection : IEnumerable<Window>
    {

        #region fields

        Workbook _workbookFilter;

        #endregion fields

        #region constructors

        public WindowsCollection()
        {
        }

        public WindowsCollection(Workbook workbook) => _workbookFilter = workbook;

        #endregion constructors

        #region properties
        #endregion properties

        #region methods

        public IEnumerator<Window> GetEnumerator()
        {

            object[,] windowsArray = (object[,])XlCall.Excel(XlCall.xlfWindows);

            int numWindows = windowsArray.GetLength(1);

            if (_workbookFilter == null)
            {
                for (int i = 0; i < numWindows; i++)
                {
                    yield return new Window((string)windowsArray[0, i]);
                }
            }
            else
            {
                for (int i = 0; i < numWindows; i++)
                {
                    //check the workbook of the window
                    string windowText = (string)windowsArray[0, i];
                    Window testWindow = new Window(windowText);

                    if (String.Equals(testWindow.WorkbookName, this._workbookFilter.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        yield return new Window(windowText);
                    }
                }
            }
            yield break;


        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion methods

    }
}
