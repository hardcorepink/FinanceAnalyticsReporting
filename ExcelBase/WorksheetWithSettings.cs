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
        public abstract void SaveClassSettings();

        public abstract string ReportSettingsAnchor
        {
            get;
        }

        //TODO have a property here for name of range to look for for settings anchor
        protected List<SettingItem> settingsList;
        private ExcelEnums.DirectionType _settingFlowDirection;

        //constructor - remember default worksheetBase constructor will be called
        public WorksheetWithSettings(ExcelEnums.DirectionType settingsFlowDirection = ExcelEnums.DirectionType.Down)
        { 
            System.Diagnostics.Debug.WriteLine("Data worksheet settings base ctor called");
            this._settingFlowDirection = settingsFlowDirection;
            this.SaveClassSettings();
        }

        public void RecalculateSettings()
        {
            //get the block as an excel reference
            object settingsBlock = this.ReturnExcelRefSettingsBlock();
            if (settingsBlock is ExcelReference sb)
            {

            }
        }

        private ExcelReference ReturnExcelRefSettingsBlock()
        {
            ExcelReference settingsAnchorBlock = base.ReturnNamedRangeRef("reportSettings");
            ExcelReference resizedSettingsBlock = null;
            if (settingsAnchorBlock != null)
            {
                ExcelReference fullSettingsBlock = base.AnchorCellToEmptySpace(settingsAnchorBlock, this._settingFlowDirection);

                switch (this._settingFlowDirection)
                {
                    case ExcelEnums.DirectionType.Down:
                    case ExcelEnums.DirectionType.Up:
                        //reference needs to be 4 columns wide 
                        if (fullSettingsBlock.ColumnLast - fullSettingsBlock.ColumnFirst != 3)
                        {
                            resizedSettingsBlock = new ExcelReference
                            (fullSettingsBlock.RowFirst, fullSettingsBlock.RowLast, fullSettingsBlock.ColumnFirst, fullSettingsBlock.ColumnFirst + 3, base.WorkSheetPtr);
                        };
                        break;

                    case ExcelEnums.DirectionType.ToLeft:
                    case ExcelEnums.DirectionType.ToRight:
                        if (fullSettingsBlock.RowLast - fullSettingsBlock.RowFirst != 3)
                        {
                            resizedSettingsBlock = new ExcelReference
                            (fullSettingsBlock.RowFirst, fullSettingsBlock.RowFirst + 3, fullSettingsBlock.ColumnFirst, fullSettingsBlock.ColumnLast, base.WorkSheetPtr);
                        };
                        break;

                }
                return resizedSettingsBlock;
            }
            else return null;

        }

        public void ReadSettingsToDictionary()
        {
            ExcelReference settingsBlock = this.ReturnExcelRefSettingsBlock();
            if (settingsBlock != null)
            {

                object settingsBlockValues = settingsBlock.GetValue();
                if (settingsBlockValues is object[,] objBlockValues)
                {
                    long rows = objBlockValues.GetLength(0);
                    this.settingsList = new List<SettingItem>();

                    //loop through - otherwise failed
                    for (long i = 0; i < rows; i++)
                    {
                        this.settingsList.Add(new SettingItem(
                            objBlockValues[i, 0].ToString(),
                            objBlockValues[i, 1].ToString(),
                            objBlockValues[i, 2].ToString(),
                            objBlockValues[i, 3].ToString()));

                        System.Diagnostics.Debug.WriteLine(settingsList[settingsList.Count - 1].ToString());
                    }
                }
                else
                {
                    //empty List of settings if failed

                    this.settingsList = new List<SettingItem>();
                }

                this.SaveClassSettings();
            }

            //TODO remove this later as this is a test line
            this.SaveIncomingSettingsToDictionary(this.settingsList);
        }

        public virtual void SaveIncomingSettingsToDictionary(List<SettingItem> incomingSettingsDictionary)
        {
            //need to loop through all settings (except for class settings as these are not shared...
            var genericSettingsIn = incomingSettingsDictionary.Where(x => String.Equals(x.SettingType, "genericSetting", StringComparison.OrdinalIgnoreCase));
            var thisGenericSettings = this.settingsList.Where(x => String.Equals(x.SettingType, "genericSetting", StringComparison.OrdinalIgnoreCase));

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

        public virtual void CommitDictionarySettingsToSheet()
        {
            //get the ExcelReference for the settingsAnchor
            ExcelReference settingsAnchor = base.ReturnNamedRangeRef("reportSettings");

            //resize the reference to fit the size of our settings list and for the number of columns required
            //TODO add switch for different settings directions
            ExcelReference newSettingsBlockAwaitingInput = new ExcelReference(settingsAnchor.RowFirst, this.settingsList.Count + settingsAnchor.RowFirst - 1, settingsAnchor.ColumnFirst, settingsAnchor.ColumnFirst + 3);

            //setup our string array - rows first
            string[,] stringArrayToSave = new string[settingsList.Count, 4];

            long arrayRowCounter = 0;
            foreach (SettingItem s in settingsList)
            {
                stringArrayToSave[arrayRowCounter, 0] = s.SettingType;
                stringArrayToSave[arrayRowCounter, 1] = s.SettingName;
                stringArrayToSave[arrayRowCounter, 2] = s.SettingValue;
                stringArrayToSave[arrayRowCounter, 3] = s.SettingSecondaryValue;
                arrayRowCounter++;
            }

            newSettingsBlockAwaitingInput.SetValue(stringArrayToSave);

        }
    }

    public class SettingItem
    {
        private string _settingName;
        private string _settingType;
        private string _settingValue;
        private string _settingSecondaryValue;

        //ctor to take all setting values, types etc.
        public SettingItem(string SettingType, string SettingName, string SettingValue, string SettingSecondaryValue)
        {
            _settingName = SettingName; _settingType = SettingType; _settingValue = SettingValue; _settingSecondaryValue = SettingSecondaryValue;
        }

        public string SettingName
        {
            get { return _settingName; }
            set { _settingName = value; }
        }

        public string SettingType
        {
            get { return _settingType; }
            set { _settingType = value; }
        }

        public string SettingValue
        {
            get { return _settingValue; }
            set { _settingValue = value; }
        }

        public string SettingSecondaryValue
        {
            get { return _settingSecondaryValue; }
            set { _settingSecondaryValue = value; }
        }

        public override string ToString()
        {
            string returnString = String.Format("SettingType = {0} | SettingName = {1} | SettingValue = {2}, SettingSecondaryValue = {3}",
                this.SettingType, this.SettingName, this.SettingValue, this.SettingSecondaryValue);

            return returnString;
        }
    }
}
