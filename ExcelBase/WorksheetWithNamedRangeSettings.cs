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
        string _settingType;            //this is the type e.g. classSetting, genericSetting, etc.
        string _settingName;            //this is the named range name
        string _settingRefText;         //the name value as defined in the name manager. e.g. if Sales is defines as 523 returns "=523"           
        object _evaluatedSettingValue;  //returns the evaluated ref text using Evaluate formula. 
        string _settingSecondaryValue;  //a secondary setting that can use for extra info, placeholders etc.
        string _settingGUISerializationString;  //gui string to define how this setting will display gui control
        

        //need a default constructor for serialization
        public NamedRangeSetting()
        {

        }


        public string SettingName
        {
            get { return _settingName; }
            set { this._settingName = value; }
        }

        public string SettingSecondaryValue
        {
            get { return _settingSecondaryValue; }
            set { _settingSecondaryValue = value; }
        }

        public string SettingGUISerializationString
        {
            get { return _settingGUISerializationString; }
            set { _settingGUISerializationString = value; }
        }

        public NamedRangeSettings ParentSettingsList
        {
            get { return this._parentListNamedRangeSettings; }  
            set { this._parentListNamedRangeSettings = value; }
        }

    }



}



