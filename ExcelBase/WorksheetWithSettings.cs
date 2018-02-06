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
                            new ExcelReference(
                                settingsBlock.RowFirst + i,
                                settingsBlock.RowFirst + i,
                                settingsBlock.ColumnFirst,
                                settingsBlock.ColumnLast),

                            objBlockValues[i, 0].ToString(),
                            objBlockValues[i, 1].ToString(),
                            objBlockValues[i, 2].ToString(),
                            objBlockValues[i, 3].ToString(),
                            objBlockValues[i, 4].ToString()
                            ));
                    }
                }
                else
                {
                    //empty List of settings if failed
                    this.listAllSettings = new List<SettingItem>();
                }

                //ok now we have all settings. In this base class we will do some sorting of the settings for the derived classes
                this.listGenericSettings = new List<SettingItem>();
                this.listGenericSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType, "genericSetting", StringComparison.OrdinalIgnoreCase)).ToList();

                this.listClassSettings = new List<SettingItem>();
                this.listClassSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType, "classSetting", StringComparison.OrdinalIgnoreCase)).ToList();

                this.listClassSettings = new List<SettingItem>();
                this.listClassSettings = this.listAllSettings.Where(s =>
                    string.Equals(s.SettingType, "calcualtedSetting", StringComparison.OrdinalIgnoreCase)).ToList();

                this.SettingsSavedToClass();
            }
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            System.Diagnostics.Debug.WriteLine($"Read settings in {elapsedMs.ToString()} ms");
        }

        public virtual void SaveIncomingSettingsToList(List<SettingItem> incomingSettingsList)
        {
            //filter out the settings, as settings in will make availabe genericSettings and calculatedSettings, but this class will only overwrite instances of genericsettings
            var genericSettingsIn = incomingSettingsList.Where(x => String.Equals(x.SettingType, "genericSetting", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(x.SettingType, "calculatedSetting", StringComparison.OrdinalIgnoreCase));

            var thisGenericSettings = this.listAllSettings.Where(x => String.Equals(x.SettingType, "genericSetting", StringComparison.OrdinalIgnoreCase));

            foreach (SettingItem inSetting in genericSettingsIn)
            {
                foreach (SettingItem currentSetting in thisGenericSettings)
                {
                    if (string.Equals(inSetting.SettingName, currentSetting.SettingName, StringComparison.OrdinalIgnoreCase))
                    {
                        currentSetting.SettingValue = inSetting.SettingValue;
                        currentSetting.SettingSecondaryValue = inSetting.SettingSecondaryValue;
                    }
                }
            }
        }

        public virtual void CommitAllSettingsToSheet()
        {

            Application.TurnScreenUpdatingOff();

            var watch = System.Diagnostics.Stopwatch.StartNew();

            List<SettingItem> commitList = this.listAllSettings.Where(s => (string.Equals(s.SettingType, "calculatedSetting", StringComparison.OrdinalIgnoreCase) == false)).ToList();

            commitList.ForEach(s => s.SaveSettingToSheet());

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            System.Diagnostics.Debug.WriteLine($"Commit settings in {elapsedMs.ToString()} ms");

            Application.TurnScreenUpdatingOn();

        }
    }

    public class SettingItem
    {
        private string _settingName;
        private string _settingType;
        private string _settingValue;
        private string _settingSecondaryValue;
        private string _settingUISerialization;

        private ExcelReference _settingExcelReference;


        //ctor to take all setting values, types etc.
        public SettingItem(ExcelReference ExcelReferenceOfSetting, string SettingType, string SettingName, string SettingValue, string SettingSecondaryValue, string SettingUISerialization)
        {
            _settingName = SettingName; _settingType = SettingType; _settingValue = SettingValue; _settingSecondaryValue = SettingSecondaryValue; _settingUISerialization = SettingUISerialization;
            _settingExcelReference = ExcelReferenceOfSetting;
        }

        public string SettingType { get => _settingType; set => _settingType = value; }
        public string SettingName { get => _settingName; set => _settingName = value; }
        public string SettingValue { get => _settingValue; set => _settingValue = value; }
        public string SettingSecondaryValue { get => _settingSecondaryValue; set => _settingSecondaryValue = value; }
        public string SettingUISerialization { get => _settingUISerialization; set => _settingUISerialization = value; }

        public override string ToString()
        {
            string returnString = String.Format("SettingType = {0} | SettingName = {1} | SettingValue = {2}, SettingSecondaryValue = {3}",
                this.SettingType, this.SettingName, this.SettingValue, this.SettingSecondaryValue);

            return returnString;
        }

        public void SaveSettingToSheet()
        {
            //save the settings array to the _settingExcelReference
            //build the array
            object testArray = this._settingExcelReference.GetValue();

            string[,] stringArray = new string[1, 5];

            stringArray[0, 0] = this.SettingType;
            stringArray[0, 1] = this.SettingName;
            stringArray[0, 2] = this.SettingValue;
            stringArray[0, 3] = this.SettingSecondaryValue;
            stringArray[0, 4] = this.SettingUISerialization;

            for (int i = 0; i < 5; i++)
            {
                if (String.Equals(stringArray[0, i], "ExcelDna.Integration.ExcelEmpty")) { stringArray[0, i] = ""; }
            }

            this._settingExcelReference.SetValue(stringArray);

        }


    }
}
