using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class NamedRange
    {

        public enum NamedRangeScope
        {
            WorkbookScoped, WorksheetScoped
        }

        #region fields

        private Worksheet _worksheet;
        private Workbook _workbook;
        private string _shortName;
        private string _fullRefName;
        private NamedRangeScope _namedRangeScope;

        #endregion fields

        #region constructors

        private NamedRange(Worksheet worksheet, string shortName)
        {
            //we are getting a worksheet scoped name range - this will get a hadle to the name
            //if it doesn't exist will return error
            

        }

        private NamedRange(Workbook workbook, string shortName)
        {
            

        }


        

        public static NamedRange ReturnExistingNamedRange()
        {
            return null;
        }

        public static NamedRange CreateNamedRange()
        {
            return null;
        }
        #endregion constructors

        #region properties

        #endregion properties





        #region methods

        public bool WorksheetNameExists()



        #endregion methods




    }
}
