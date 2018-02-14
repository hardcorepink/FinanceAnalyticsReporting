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
                
        }



    }

    public class WindowsCollection : IEnumerable<Window>
    {
        Workbook _workbookFilter;

        public WindowsCollection()
        {
        }

        public WindowsCollection(Workbook workbook) => _workbookFilter = workbook;
        

        public IEnumerator<Window> GetEnumerator()
        {

            object[,] windowsArray = (object[,])XlCall.Excel(XlCall.xlfWindows);

            int numWindows = windowsArray.GetLength(1);

            if (_workbookFilter == null)
            {
                for(int i = 0; i < numWindows; i++)
                {
                    yield return new Window((string)windowsArray[0, i]);
                }
            }
            else
            {
                for(int i = 0; i < numWindows; i++)
                {
                    //check the workbook of the window
                    string windowText = (string)windowsArray[0, i];
                    string workbookText = (string)XlCall.Excel(XlCall.xlfGetWindow, 31, windowText);
                    
                    int charLocation = workbookText.IndexOf(":", StringComparison.Ordinal);
                    if (charLocation > 0)
                    {
                        workbookText = workbookText.Substring(0, charLocation);
                    }

                    if(String.Equals(workbookText, this._workbookFilter.Name, StringComparison.OrdinalIgnoreCase))
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
    }
}
