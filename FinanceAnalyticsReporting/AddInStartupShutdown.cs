using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelBase;
using System.Diagnostics;

namespace FinanceAnalyticsReporting
{
    public class AddInStartupShutdown : ExcelDna.Integration.IExcelAddIn
    {

        public void AutoOpen()
        {
            //register our event on worksheet change
            XlCall.Excel(XlCall.xlcOnSheet, null, "WorksheetSelectionChanged");
        }

        public void AutoClose()
        {

        }

    }

    public static class EventCallbacks
    {
        [ExcelCommand()]
        public static void WorksheetSelectionChanged()
        {
            ExcelBase.Worksheet testWS;
            //build a new worksheet based on our ExcelBase Assembly
            try
            {
                testWS = new Worksheet();
            }
            catch
            {
                XlCall.Excel(XlCall.xlcMessage, true, string.Format("New selection is not a worksheet"));
                return;
            }


            //output some statistics on the selected worksheet
            Debug.WriteLine("Short Worksheet Name : " + testWS.ShortWorksheetName.ToString());
            Debug.WriteLine("Workseet Ptr : " + testWS.WorkSheetPtr);
            Debug.WriteLine("Workbook Name : " + testWS.WorkbookName);
            Debug.WriteLine("Full WS Name: " + testWS.FullWorksheetName);
            Debug.WriteLine("WB Details (Name): " + testWS.ParentWorkbook.Name);

            string worksheetName = testWS.ShortWorksheetName;

            XlCall.Excel(XlCall.xlcMessage, true, string.Format("ExcelDNA loaded active sheet is {0}", worksheetName));

            //System.Windows.MessageBox.Show($"Worksheet Name is : {worksheetName}");

        }
    }
}
