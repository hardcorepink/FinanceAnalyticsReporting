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
    public static class ClassTesting
    {
                       
        [ExcelCommand(MenuName = "Hello", MenuText = "OpenWorkbook")]
        public static void OpenWorkbook()
        {
            //get activeSheet
            ExcelBase.Worksheet activeWS = new ExcelBase.Worksheet();
            ExcelReference newExcelRef = new ExcelReference(0, 0, 0, 0, activeWS.WorkSheetPtr);
            string workbookToTest = (string)newExcelRef.GetValue();

            try
            {
                Workbook newWorkbook = new Workbook(workbookToTest);
                System.Windows.MessageBox.Show("Opened workbook: " + newWorkbook.Name);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "Create new workbook")]
        public static void CreateNewWorkbook()
        {

            Workbook newWB = new Workbook(true);
        }

        [ExcelCommand(MenuName = "Hello", MenuText ="Hide all windows belonging to active workbook")]
        public static void HideAllWindows()
        {
            Workbook newWB = new Workbook(false);
            newWB.HideAllWorkbookWindows();

        }


       
    }
}
