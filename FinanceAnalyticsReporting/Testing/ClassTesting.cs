﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelBase;
using System.Diagnostics;
using FinanceAnalyticsReporting.ExcelWorksheetTypes;

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

        [ExcelCommand(MenuName = "Hello", MenuText = "Hide all windows belonging to active workbook")]
        public static void HideAllWindows()
        {
            Workbook newWB = new Workbook(false);
            newWB.HideAllWorkbookWindows();

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "Create new worksheet from active sheet")]
        public static void CreateNewWorksheetObject()
        {

            testWS = new Worksheet();
            // Debug.WriteLine("SheetRef: " + newWS.SheetRef.ToString());
            Debug.WriteLine("Short Worksheet Name : " + testWS.ShortWorksheetName.ToString());
            Debug.WriteLine("Workseet Ptr : " + testWS.WorkSheetPtr);
            Debug.WriteLine("Workbook Name : " + testWS.WorkbookName);
            Debug.WriteLine("Full WS Name: " + testWS.FullWorksheetName);
        }

        public static ExcelBase.Worksheet testWS;

        [ExcelCommand(MenuName = "Hello", MenuText = "TestMovingWS")]
        public static void TestMovingWS()
        {

            // Debug.WriteLine("SheetRef: " + newWS.SheetRef.ToString());
            Debug.WriteLine("Short Worksheet Name : " + testWS.ShortWorksheetName.ToString());
            Debug.WriteLine("Workseet Ptr : " + testWS.WorkSheetPtr);
            Debug.WriteLine("Workbook Name : " + testWS.WorkbookName);
            Debug.WriteLine("Full WS Name: " + testWS.FullWorksheetName);
            Debug.WriteLine("WB Details (Name): " + testWS.ParentWorkbook.Name);

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "TestWorksheetAlive")]
        public static void TestWorksheetAlive()
        {
            System.Windows.MessageBox.Show(testWS.IsPointerStillValid().ToString());

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "TestWorksheetIndexer")]
        public static void TestWorksheetIndexer()
        {
            Workbook newWB = new Workbook(false);
            Worksheet tryWS = newWB.Worksheets["Hello"];
            if (tryWS != null)
            {
                string wsName = tryWS.ShortWorksheetName;
                System.Windows.MessageBox.Show(tryWS.ShortWorksheetName);
            }
            else
            {
                System.Windows.MessageBox.Show("Could not find Hello sheet in active workbook");
            }
                
        }

        [ExcelCommand(MenuName = "Hello", MenuText = "TestFullWorksheetConstrcutor")]
        public static void TestFullWorksheetConstrcutor()
        {

            var worksheetString = "[Book4]Hello";
            var newWS = new Worksheet(worksheetString);
                        
        }

        [ExcelCommand(MenuName = "Hello", MenuText = "TestWorksheetIterator")]
        public static void TestWorksheetIterator()
        {

            var activeWB = new Workbook(false);

            foreach(Worksheet ws in activeWB.Worksheets)
            {
                System.Windows.MessageBox.Show(ws.FullWorksheetName);
            }

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "Evalate formula under selection")]
        public static void EvaluateFormulaInSelection()
        {

            //first get the active cell
            ExcelReference activeCell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);

            //get the formula of the activeCell
            object activeCellFormula = XlCall.Excel(XlCall.xlfGetFormula, activeCell);
            string cellFormulaString = (string)activeCellFormula;

            //evaluate the formula and display result as a messageBox
            object Result = XlCall.Excel(XlCall.xlfGetCell, 6, activeCell);

            string evalResult = (XlCall.Excel(XlCall.xlfEvaluate, Result)).ToString();

            System.Windows.MessageBox.Show(evalResult);


        }


        [ExcelCommand(MenuName = "Hello", MenuText = "Read Settings then save settings back to sheet")]
        public static void ReadThenSaveSettings()
        {

            //first get instance of the activeSheet as a reportsheet
            ReportWorksheet newReportSheet = new ReportWorksheet();

            //now commit the read settings back to sheet
            newReportSheet.CommitAllSettingsToSheet();

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "Reload Report Worksheet")]
        public static void ReloadReportWorksheet()
        {

            //first get instance of the activeSheet as a reportsheet
            ReportWorksheet newReportSheet = new ReportWorksheet();

            //now commit the read settings back to sheet
            newReportSheet.ReloadReportWorksheet();

        }


    }
}
