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

        #region properties

        public static bool ScreenUpdating
        {
            get { return (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 40); }
            set { XlCall.Excel(XlCall.xlcEcho, value); }
        }

        public static WorkbookCollection Workbooks
        {
            get { return new WorkbookCollection(); }
        }

        #endregion properties               

    }



}









