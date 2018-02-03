using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelBase;
using System.Diagnostics;

namespace FinanceAnalyticsReporting.Application
{
    public class AddInStartupShutdown : ExcelDna.Integration.IExcelAddIn
    {

        public void AutoOpen()
        {
            //register our event on worksheet change
            XlCall.Excel(XlCall.xlcOnSheet, null, "WorksheetChangedApp");

            StaticAppState.BuildDictionaryOfWorksheetTypes();
        }

        public void AutoClose()
        {

        }

    }
        
}
