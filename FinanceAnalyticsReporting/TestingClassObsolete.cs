/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelBase;


namespace FinanceAnalyticsReporting
{
    public static class TestingClass
    {

        [ExcelCommand(MenuName = "Hello", MenuText = "Hello")]
        public static void TestSettingsRange()
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            //get handle to active sheet
            var res = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string sheetName = (string)res;

            //now get the worksheet short name
            int pos = sheetName.IndexOf(']');
            string worksheetName = sheetName.Substring(pos + 1, sheetName.Length - pos - 1);

            //now look for a reference of the "repoortSettings"
            string searchNamedRange = string.Format("'{0}'!reportSettings", worksheetName);

            var origSelection = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;

            object value = XlCall.Excel(XlCall.xlfEvaluate, searchNamedRange);
            if (value is ExcelReference)
            {
                ExcelReference ExcelRefOut = (ExcelReference)value;
                XlCall.Excel(XlCall.xlcSelect, ExcelRefOut);
                ExcelReference nextCell = ExcelRefOut.End(BaseWorksheet.DirectionType.Down);
                string outString = nextCell.GetValue().ToString();
                System.Windows.MessageBox.Show(outString);
            }


            XlCall.Excel(XlCall.xlcSelect, origSelection);

            XlCall.Excel(XlCall.xlcEcho, true);

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "get Active Sheet Name")]
        public static void ReturnSheetName()
        {
            //gets reference to activesheet as a sheetID
            var sheetID = XlCall.Excel(XlCall.xlSheetId);


            ExcelReference newExcelRef = (ExcelReference)sheetID;
            System.IntPtr newIntPtr = newExcelRef.SheetId;
            System.Windows.MessageBox.Show(newIntPtr.ToString());

            BaseWorksheet newWSBase = new BaseWorksheet(newIntPtr);

            System.Windows.MessageBox.Show(newWSBase.WorksheetName);

        }


        [ExcelCommand(MenuName = "Hello", MenuText = "test Settings Workbook")]
        public static void TestSettingsWorkbook()
        {
            WorksheetWithSettings newWSSettingsBase = new ActiveFormDataProvider();
            newWSSettingsBase.ReadSettingsToDictionary();

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "test alter settings and save")]
        public static void TestAlterSettingsAndSavek()
        {
            WorksheetWithSettings newWSSettingsBase = new ActiveFormDataProvider();
            newWSSettingsBase.ReadSettingsToDictionary();

            //alter the settings
            List<SettingItem> alterSettings = new List<SettingItem>();
            alterSettings.Add(new SettingItem("genericSetting", "testname", "testValue1", "testValue2"));

            newWSSettingsBase.SaveIncomingSettingsToDictionary(alterSettings);
            newWSSettingsBase.CommitDictionarySettingsToSheet();


        }

        [ExcelCommand(MenuName = "Hello", MenuText = "Go Left, Right, Up, Down etc.")]
        public static void SelectAnchor()
        {
            BaseWorksheet newWS = new BaseWorksheet();

            ExcelReference directionValueRange = newWS.ReturnNamedRangeRef("testValue");
            int directionValue = Convert.ToInt32(directionValueRange.GetValue());

            var selection = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;


            if (typeof(BaseWorksheet.DirectionType).IsEnumDefined(directionValue))
            {
                ExcelReference newExcelRef = newWS.AnchorCellToEmptySpace(selection, (BaseWorksheet.DirectionType)directionValue);
                XlCall.Excel(XlCall.xlcSelect, newExcelRef);
            }

            //select the reference

        }

        [ExcelCommand(MenuName = "Hello", MenuText = "ProtectActiveSheet")]
        public static void ProtectActiveSheet()
        {
            BaseWorksheet newWs = new BaseWorksheet();
            XlCall.Excel(XlCall.xlcProtectDocument, true, false, ExcelMissing.Value, false, true);
        }




        public static ExcelReference End(this ExcelReference reference, BaseWorksheet.DirectionType direction)
        {
            ExcelReference end = null;

            // myReference is selected now...
            XlCall.Excel(XlCall.xlcSelectEnd, (int)direction);

            var selection = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;
            var row = selection.RowFirst;
            var col = selection.ColumnFirst;

            end = new ExcelReference(row, row, col, col, selection.SheetId);

            return end;
        }
    }
}
*/