using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelBase;

namespace FinanceAnalyticsReporting
{
    public class AddInStartupShutdown : ExcelDna.Integration.IExcelAddIn
    {

        public void AutoOpen()
        {

            //register our event on worksheet change
            XlCall.Excel(XlCall.xlcOnSheet, null, "WorksheetChanged");
        }

        public void AutoClose()
        {

        }


    }

    public static class EventCallbacks
    {
        [ExcelCommand()]
        public static void WorksheetChanged()
        {

            System.Windows.MessageBox.Show("Worksheet chnaged");
            //build a new worksheet based on our ExcelBase Assembly
            ExcelBase.Worksheet activeSheet = new Worksheet();
            string worksheetName = activeSheet.WorksheetName;
            System.Windows.MessageBox.Show($"Worksheet Name is : {worksheetName}");



        }
    }
}
