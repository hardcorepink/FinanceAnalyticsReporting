using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;


namespace ExcelBase
{
    public class WorksheetWithNamedRangeSettings : Worksheet

    {

        private NamedRangeSettings _settingsList;

        /// <summary>
        /// Constructor for active sheet
        /// </summary>
        public WorksheetWithNamedRangeSettings() : base()
        {

        }


        public NamedRangeSettings SettingsList
        {
            get { return _settingsList; }
            set { _settingsList = value; }
        }


        public void BinarySerializeSettingsToA2()
        {
            string outString;
            using (MemoryStream ms = new MemoryStream())
            {
                new BinaryFormatter().Serialize(ms, this._settingsList);
                outString = Convert.ToBase64String(ms.ToArray());
            }

            ExcelReference newExcelRef = new ExcelReference(1, 1, 0, 0, this.WorkSheetPtr);
            newExcelRef.SetValue(outString);
        }

        public void BinaryDeSerializeA2ToSettings()
        {
            ExcelReference newExcelRef = new ExcelReference(1, 1, 0, 0, this.WorkSheetPtr);
            string xmlString = newExcelRef.GetValue().ToString();
            byte[] bytes = Convert.FromBase64String(xmlString);
            using (MemoryStream ms = new MemoryStream(bytes, 0, bytes.Length))
            {
                ms.Write(bytes, 0, bytes.Length);
                ms.Position = 0;
                object DeserializedObject = new BinaryFormatter().Deserialize(ms);
                this._settingsList = (NamedRangeSettings)DeserializedObject;
                this._settingsList.ParentSheet = this;

            }
        }

    }


    public class NamedRangeSettings
    {
        private List<NamedRangeSetting> _listNamedRangeSettings;

        [NonSerialized]
        private WorksheetWithNamedRangeSettings _parentSheet;

        public List<NamedRangeSetting> ListNamedRangeSettings
        {
            get { return _listNamedRangeSettings; }
            set { }
        }


        public NamedRangeSettings()
        {
            this._listNamedRangeSettings = new List<NamedRangeSetting>();
        }

        public void AddSetting(NamedRangeSetting NamedRangeSetting)
        {
            this._listNamedRangeSettings.Add(NamedRangeSetting);
            NamedRangeSetting.ParentSettingsList = this;
        }

        public WorksheetWithNamedRangeSettings ParentSheet
        {
            get
            {
                return this._parentSheet;
            }
            set
            {
                this._parentSheet = value;
                foreach (NamedRangeSetting nrs in this.ListNamedRangeSettings) { nrs.ParentSettingsList = this; }
            }
        }

    }

    public class NamedRangeSetting
    {
        NamedRangeSettings _parentListNamedRangeSettings;
        private string _settingType;                    //this is the type e.g. classSetting, genericSetting, etc.
        private string _settingName;                    //this is the named range name
        private string _settingNameDefinition;                 //the definition of the name - e.g. "=A1", "=523" etc.      
        private object _evaluatedSettingValue;          //returns the evaluated ref text using Evaluate formula. 
        private string _settingSecondaryValue;          //a secondary setting that can use for extra info, placeholders etc.
        private string _settingGUISerializationString;  //gui string to define how this setting will display gui control
        private bool _settingSharedOut;
        private bool _settingCanBeOverwritten;

        private string _settingFullName;


        public string SettingType
        {
            get => _settingType;
            set => _settingType = value;
        }

        public string SettingName
        {
            get => _settingName;

            set
            {
                //try and define a new setting



                //try and delete the old setting name
                try
                {
                    XlCall.Excel(XlCall.xlcDeleteName, this.SettingFullName);
                }
                catch { }

            }
        }

        public string SettingFullName
        {
            get
            {
                string wsName = this.ParentSettingsList.ParentSheet.FullWorksheetName;
                return wsName + "!" + this._settingName;
            }

        }

        public string SettingRefText
        {
            get
            {
                return (string)XlCall.Excel(XlCall.xlfGetName, this.SettingFullName);
            }
            set
            {
                XlCall.Excel(XlCall.xlfSetName, this.SettingFullName, value);
            }
        }


        public string SettingSecondaryValue
        {
            get => _settingSecondaryValue;
            set => _settingSecondaryValue = value;
        }

        public string SettingGUISerializationString
        {
            get => _settingGUISerializationString;
            set => _settingGUISerializationString = value;
        }

        public NamedRangeSettings ParentSettingsList
        {
            get => this._parentListNamedRangeSettings;
            set => this._parentListNamedRangeSettings = value;
        }

    }



}



