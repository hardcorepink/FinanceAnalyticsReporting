using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public abstract class WorksheetWithSettings : Worksheet
    {

        #region fields
        //All lists of settings here. Class settings can have lists of settings, as they will not need to be 
        protected List<SettingItem> listClassSettings;
        protected List<SettingItem> listGenericSettings;
        protected List<SettingItem> listCalculatedSettings;
        protected List<SettingItem> listAllSettings;

        //this will be triggered to the derived class when there are new settings. I.e. either reloaded from the sheet or 
        //imported from another report
        public abstract void SettingsSavedToClass();

        //setting fill direction always looks down     
        private ExcelEnums.DirectionType _settingFlowDirection = ExcelEnums.DirectionType.Down;


        #endregion fields

        //default constuctor here - calls base constuctor to look construct worksheet from active sheet
        public WorksheetWithSettings() : base()
        {

        }

        private ExcelReference ReturnExcelRefSettingsBlock()
        {
            ExcelReference settingsAnchorBlock = base.ReturnNamedRangeRef("settingsAnchor");
            ExcelReference resizedSettingsBlock = null;
            if (settingsAnchorBlock != null)
            {
                ExcelReference fullSettingsBlock = base.AnchorCellToEmptySpace(settingsAnchorBlock, this._settingFlowDirection);

                switch (this._settingFlowDirection)
                {
                    case ExcelEnums.DirectionType.Down:
                    case ExcelEnums.DirectionType.Up:
                        //reference needs to be 4 columns wide 
                        if (fullSettingsBlock.ColumnLast - fullSettingsBlock.ColumnFirst != 4)
                        {
                            resizedSettingsBlock = new ExcelReference
                            (fullSettingsBlock.RowFirst, fullSettingsBlock.RowLast, fullSettingsBlock.ColumnFirst, fullSettingsBlock.ColumnFirst + 4, base.WorkSheetPtr);
                        };
                        break;

                    case ExcelEnums.DirectionType.ToLeft:
                    case ExcelEnums.DirectionType.ToRight:
                        if (fullSettingsBlock.RowLast - fullSettingsBlock.RowFirst != 4)
                        {
                            resizedSettingsBlock = new ExcelReference
                            (fullSettingsBlock.RowFirst, fullSettingsBlock.RowFirst + 4, fullSettingsBlock.ColumnFirst, fullSettingsBlock.ColumnLast, base.WorkSheetPtr);
                        };
                        break;

                }
                return resizedSettingsBlock;
            }
            else return null;

        }

        /// <summary>
        /// This method reads a settings from excel sheet into various lists of type settingItem.
        /// </summary>
        public void ReadSettingsToList()
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            ExcelReference settingsBlock = this.ReturnExcelRefSettingsBlock();

            if (settingsBlock != null)
            {

                object settingsBlockValues = settingsBlock.GetValue();
                if (settingsBlockValues is object[,] objBlockValues)
                {
                    int rows = objBlockValues.GetLength(0);
                    this.listAllSettings = new List<SettingItem>();

                    //loop through - otherwise failed
                    for (int i = 0; i < rows; i++)
                    {
                        this.listAllSettings.Add(new SettingItem(
                            new ExcelReference(settingsBlock.RowFirst + i, settingsBlock.RowFirst + i, settingsBlock.ColumnFirst, settingsBlock.ColumnLast),
                            objBlockValues[i, 0], objBlockValues[i, 1], objBlockValues[i, 2], objBlockValues[i, 3], objBlockValues[i, 4]
                            ));
                    }
                }
                else
                {
                    //empty List of settings if failed
                    this.listAllSettings = new List<SettingItem>();
                }

                //ok now we have all settings. In this base class we will do some sorting of the settings for the derived classes
                this.listGenericSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType.ToString(), "genericSetting", StringComparison.OrdinalIgnoreCase)).ToList();

                this.listClassSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType.ToString(), "classSetting", StringComparison.OrdinalIgnoreCase)).ToList();
                                
                this.listCalculatedSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType.ToString(), "calcualtedSetting", StringComparison.OrdinalIgnoreCase)).ToList();

                this.SettingsSavedToClass();
            }
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            System.Diagnostics.Debug.WriteLine($"Read settings in {elapsedMs.ToString()} ms");
        }

        public virtual void SaveIncomingSettingsToList(List<SettingItem> incomingSettingsList)
        {
            //filter out the settings, as settings in will make availabe genericSettings and calculatedSettings, but this class will only overwrite instances of genericsettings
            var genericSettingsIn = incomingSettingsList.Where(x => String.Equals(x.SettingType.ToString(), "genericSetting", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(x.SettingType.ToString(), "calculatedSetting", StringComparison.OrdinalIgnoreCase));

            var thisGenericSettings = this.listAllSettings.Where(x => String.Equals(x.SettingType.ToString(), "genericSetting", StringComparison.OrdinalIgnoreCase));

            foreach (SettingItem inSetting in genericSettingsIn)
            {
                foreach (SettingItem currentSetting in thisGenericSettings)
                {
                    if (string.Equals(inSetting.SettingName.ToString(), currentSetting.SettingName.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        currentSetting.SettingValue = inSetting.SettingValue;
                        currentSetting.SettingSecondaryValue = inSetting.SettingSecondaryValue;
                    }
                }
            }
        }

        public virtual void CommitAllSettingsToSheet()
        {

            ExcelApplication.ScreenUpdating = false;

            var watch = System.Diagnostics.Stopwatch.StartNew();

            List<SettingItem> commitList = this.listAllSettings.Where(s => (string.Equals(s.SettingType.ToString(), "calculatedSetting", StringComparison.OrdinalIgnoreCase) == false)).ToList();

            commitList.ForEach(s => s.SaveSettingToSheet());

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            System.Diagnostics.Debug.WriteLine($"Commit settings in {elapsedMs.ToString()} ms");

            ExcelApplication.ScreenUpdating = true;

        }
    }

    public class SettingItem
    {
        private object _settingType;
        private object _settingName;
        private object _settingValue;
        private object _settingSecondaryValue;
        private object _settingUISerialization;

        private ExcelReference _settingExcelReference;


        //ctor to take all setting values, types etc.
        public SettingItem(ExcelReference ExcelReferenceOfSetting, object SettingType, object SettingName, object SettingValue, object SettingSecondaryValue, object SettingUISerialization)
        {
            _settingName = SettingName; _settingType = SettingType; _settingValue = SettingValue; _settingSecondaryValue = SettingSecondaryValue; _settingUISerialization = SettingUISerialization;
            _settingExcelReference = ExcelReferenceOfSetting;
        }

        public object SettingType { get => _settingType; set => _settingType = value; }
        public object SettingName { get => _settingName; set => _settingName = value; }
        public object SettingValue { get => _settingValue; set => _settingValue = value; }
        public object SettingSecondaryValue { get => _settingSecondaryValue; set => _settingSecondaryValue = value; }
        public object SettingUISerialization { get => _settingUISerialization; set => _settingUISerialization = value; }


        public void SaveSettingToSheet()
        {
            //save the settings array to the _settingExcelReference
            //build the array
            object testArray = this._settingExcelReference.GetValue();

            object[,] outArray = new object[1, 5];

            outArray[0, 0] = this.SettingType;
            outArray[0, 1] = this.SettingName;
            outArray[0, 2] = this.SettingValue;
            outArray[0, 3] = this.SettingSecondaryValue;
            outArray[0, 4] = this.SettingUISerialization;

            this._settingExcelReference.SetValue(outArray);

        }

        public override string ToString()
        {
            string returnString = String.Format("SettingType = {0} | SettingName = {1} | SettingValue = {2}, SettingSecondaryValue = {3}",
                this.SettingType.ToString(), this.SettingName.ToString(), this.SettingValue.ToString(), this.SettingSecondaryValue.ToString());

            return returnString;
        }

    }
}
