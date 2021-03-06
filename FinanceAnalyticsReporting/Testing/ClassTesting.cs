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



        [ExcelCommand(MenuName = "Hello", MenuText = "TestFullWorksheetConstrcutor")]
        public static void TestFullWorksheetConstrcutor()
        {

            var worksheetString = "[Book4]Hello";
            var newWS = new Worksheet(worksheetString);

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

        [ExcelCommand(MenuName = "Hello", MenuText = "Evaluate Test")]
        public static void EvaluateSheet1Test()
        {
            object result3;
            object result8;
            object result1 = XlCall.Excel(XlCall.xlfGetName, @"'Sheet1'!Test");

            object result = XlCall.Excel(XlCall.xlfEvaluate, @"'Sheet1'!Test");


            string newResult1String = result1.ToString().Substring(1);

            object anotherResult = (XlCall.Excel(XlCall.xlfTextref, newResult1String, false)).ToString();

            result3 = XlCall.Excel(XlCall.xlfEvaluate, newResult1String);


            result3 = XlCall.Excel(XlCall.xlfEvaluate, result);
            result8 = ((ExcelReference)result3).GetValue();


            string outResult = $"Straight Get Name Result: {result1.ToString()} {Environment.NewLine}" +
                $"Type {result1.GetType().Name} {Environment.NewLine} {Environment.NewLine}" +
                $"Evaluate Result: {result.ToString()} {Environment.NewLine}" +
                $"Type {result.GetType().Name} {Environment.NewLine} {Environment.NewLine}" +
                $"Evaluate the Get Name - Result: {result8.ToString()} {Environment.NewLine}" +
                $"Type {result3.GetType().Name} {Environment.NewLine} {Environment.NewLine}";

            System.Windows.MessageBox.Show(outResult);


        }


        [ExcelCommand(MenuName = "Report Settings", MenuText = "putInLongStringFromA1")]
        public static void PutInLongStringFromA1()
        {
            Worksheet newWS = new Worksheet();

            ExcelReference A1Ref = new ExcelReference(0, 0, 0, 0, newWS.WorkSheetPtr);
            ExcelReference A2Ref = new ExcelReference(0, 0, 1, 1, newWS.WorkSheetPtr);

            int A1RefInt = Convert.ToInt16(A1Ref.GetValue());

            string newString = new string('a', A1RefInt);

            A2Ref.SetValue(newString);
        }

        [ExcelCommand(MenuName = "Report Settings", MenuText = "BinarySerializeToA2")]
        public static void BinarySerializeToA2()
        {
            ExcelBase.WorksheetWithNamedRangeSettings newReportWS = new ExcelBase.WorksheetWithNamedRangeSettings();

            for (int i = 0; i < 20; i++)
            {
                NamedRangeSetting nrs = new NamedRangeSetting
                {
                    SettingName = $"Hello {i.ToString()}",
                    SettingSecondaryValue = $"Hey {i.ToString()}"
                };

                newReportWS.SettingsList.AddSetting(nrs);
            }

            var watch = System.Diagnostics.Stopwatch.StartNew();

            //time this part
            newReportWS.BinarySerializeSettingsToA2();

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Debug.WriteLine($"Serialized Binary in {elapsedMs.ToString()} ms");
        }

        [ExcelCommand(MenuName = "Report Settings", MenuText = "BinaryDeSerializeFromA2")]
        public static void BinaryDeSerializeFromA2()
        {
            ExcelBase.WorksheetWithNamedRangeSettings newReportWS = new ExcelBase.WorksheetWithNamedRangeSettings();

            var watch = System.Diagnostics.Stopwatch.StartNew();

            newReportWS.BinaryDeSerializeA2ToSettings();

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Debug.WriteLine($"DeSerialized Binary in {elapsedMs.ToString()} ms");

        }

        [ExcelCommand(MenuName = "Report Settings", MenuText = "OpenWorkbookInA1")]
        public static void OpenWorkbookInA1()
        {
            Worksheet newWS = new Worksheet();
            ExcelReference excelRef = new ExcelReference(0, 0, 0, 0, newWS.WorkSheetPtr);
            string val = (string)excelRef.GetValue();
            System.IO.FileInfo newFI = new System.IO.FileInfo(val);
            Workbook newWB = ExcelApplication.Workbooks[newFI];

            try
            {
                newWB = ExcelApplication.Workbooks.Open(val);
                System.Windows.MessageBox.Show($"New workbook = {newWB.Name}");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Could not open : {ex.Message}");
            }





        }

        [ExcelCommand(MenuName = "Report Settings", MenuText = "GetWorkbookStatusInA1")]
        public static void GetWorkbookStatusInA1()
        {



        }

        [ExcelCommand(MenuName = "Report Settings", MenuText = "loopThroughWorkbooks")]
        public static void TestWokbooksLooper()
        {
            string output = "";

            foreach (Workbook wb in ExcelApplication.Workbooks)
            {
                output = output + wb.Name + Environment.NewLine;
            }

            System.Windows.MessageBox.Show(output);
        }



        [ExcelCommand(MenuName = "Big Data", MenuText = "Write Big Data To Excel")]
        unsafe public static void NewBinaryName()
        {
            //here is the array of bytes we are going to save            
            var sevenItems = new byte[] { 0x5A, 0x5A, 0x5A };

            //now try passing the byte array and name (string) to create binary 
            object result = XlCall.Excel(XlCall.xlDefineBinaryName, "Hello", sevenItems);

        }

        [ExcelCommand(MenuName = "Big Data", MenuText = "Get Big Data To Excel")]
        unsafe public static void GetBinaryData()
        {

            byte[] result = (byte[])XlCall.Excel(XlCall.xlGetBinaryName, "Hello");

            string resultString = System.Text.Encoding.UTF8.GetString(result);

            System.Windows.Forms.MessageBox.Show(resultString);


        }

        [ExcelCommand(MenuName = "Names", MenuText = "Get List of Names")]
        unsafe public static void GetListofNames()
        {

            //get values of cell a1 in active sheet

            //no marshal of paramaters required


            Worksheet activeSheet = new Worksheet();
            string valueOfFirstCell = (string)activeSheet.Range["A1"].GetValue();

            Object listOfNames = XlCall.Excel(XlCall.xlfNames, valueOfFirstCell);

        }

        [ExcelCommand(MenuName = "Names", MenuText = "Get List of WorksheetNames")]
        unsafe public static void GetListOfWorksheetNames()
        {

            //get values of cell a1 in active sheet

            //no marshal of paramaters required


            Worksheet activeSheet = new Worksheet();

            System.Windows.Forms.MessageBox.Show(activeSheet.Names.ToString());

        }

        [ExcelCommand(MenuName = "Names", MenuText = "Get List of WorkbookNames")]
        public static void GetListOfWorkbookNames()
        {
            //get values of cell a1 in active sheet

            //no marshal of paramaters required
            Workbook activeWorkbook = new Workbook();

            System.Windows.Forms.MessageBox.Show(activeWorkbook.Names.ToString());

        }

        [ExcelCommand(MenuName = "Names", MenuText = "Test Add Name")]
        public static void DefineName()
        {

            Worksheet activeWS = new Worksheet();

            string refersTo = (string)activeWS.Range["A1"].GetValue();

            Workbook activeWorkbook = new Workbook();

            NamedRange newName = activeWorkbook.Names.Add("testNewName", refersTo, false);

        }

        [ExcelCommand(MenuName = "Names", MenuText = "AddWorksheetName")]
        public static void DAddWorksheetName()
        {

            Worksheet activeWS = new Worksheet();

            string refersTo = (string)activeWS.Range["A1"].GetValue();

            //now define a new Name

            NamedRange newName = activeWS.Names.Add("testWSName", refersTo, false);

        }

        [ExcelCommand(MenuName = "Names", MenuText = "DeleteWorksheetNameInA1")]
        public static void DeleteWorksheetNameCalledTesting()
        {

            Worksheet activeWS = new Worksheet();
            NamedRange nr = activeWS.Names["testing"];
            nr.Delete();

        }

        [ExcelCommand(MenuName = "Names", MenuText = "DeleteWorkbookNameInA1")]
        public static void DeleteWorkbookNameCalledTesting()
        {
            try
            {
                string NameToDelete = (string)(new Worksheet()).Range["A1"].GetValue();
                XlCall.Excel(XlCall.xlcDeleteName, NameToDelete);
            }
            catch { }
        }


    }
}
