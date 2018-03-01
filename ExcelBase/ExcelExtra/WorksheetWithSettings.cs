using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBase.ExcelExtra
{
    public class WorksheetWithSettings : Worksheet
    {

        private SettingsCollection _settingsCollectionInstance;

        public WorksheetWithSettings() : base()
        {
            //TODO need constructor for other ways of accessing sheets
            this._settingsCollectionInstance = new SettingsCollection(this);
        }

        public void CommitSettingsToBinaryName()
        {
            byte[] byteArrayOfSettings;

            //remember can only do this for the active sheet
            this.Activate();

            //serialize our settings collection to a binary string
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())
            {
                bf.Serialize(ms, this._settingsCollectionInstance);
                byteArrayOfSettings = ms.ToArray();
            }

            //now try passing the byte array and name (string) to create binary 
            object result = XlCall.Excel(XlCall.xlDefineBinaryName, "Settings", byteArrayOfSettings);

        }

        public void ReadSettingsFromBinaryName()
        {

            //remember can only do this for the active sheet
            this.Activate();
            byte[] settingsResult = (byte[])XlCall.Excel(XlCall.xlGetBinaryName, "Settings");

            //string resultString = System.Text.Encoding.UTF8.GetString(result);            
            MemoryStream memStream = new MemoryStream();
            BinaryFormatter binForm = new BinaryFormatter();
            memStream.Write(settingsResult, 0, settingsResult.Length);
            memStream.Seek(0, SeekOrigin.Begin);

            SettingsCollection obj = (SettingsCollection)binForm.Deserialize(memStream);

        }
        
    }

    public class Setting
    {

        #region fields

        private string _settingName;

        private bool _settingExported;
        private bool _settingCanBeOverwritten;
        private bool _settingExposedAsWorksheetName;
        private bool _isClassProperty;

        private string _settingValueString;
        private string _settingSecondaryValueString;
        private object _additionalTagObject;
        private int _settingTypeIdentifier;

        private string _settingXMLSerialization;

        private SettingsCollection _parentsettingsCollection;

        #endregion fields

        #region constructors

        public Setting(string settingName, SettingsCollection ParentSettingsCollection)
        {
            this._settingName = settingName;
            this._parentsettingsCollection = ParentSettingsCollection;
        }

        #endregion constructors


        #region properties

        public string Name
        {
            get { return _settingName; }
            //cannot change name as it is the key to the settingsCollection
        }

        public bool SettingExported
        {
            get => _settingExported;
            set => _settingExported = value;
        }

        public bool SettingCanBeOverwritten
        {
            get => _settingCanBeOverwritten;
            set => _settingCanBeOverwritten = value;
        }

        public bool SettingExposedAsWorksheetName
        {
            get => _settingExposedAsWorksheetName;
            set
            {
                //TODO need to define or delete a name here//
                if(value == true)
                {
                    this._parentsettingsCollection.BaseWorksheetWithSettings.Names.Add(this.Name, this.SettingValueString, false);
                }
                else
                {

                }
            }
        }

        public bool IsClassProperty
        {
            get => _isClassProperty;
            set => _isClassProperty = value;
        }


        public string SettingValueString
        {
            get => _settingValueString;
            set => _settingValueString = value;
        }


        public string SettingSecondaryValueString
        {
            get => _settingSecondaryValueString;
            set => _settingSecondaryValueString = value;
        }

        public object AdditionalTagObject
        {
            get => _additionalTagObject;
            set => _additionalTagObject = value;
        }

        public int SettingTypeIdentifier
        {
            get => _settingTypeIdentifier;
            set => _settingTypeIdentifier = value;
        }

        public string SettingXMLSerialization
        {
            get => _settingXMLSerialization;
            set => _settingXMLSerialization = value;
        }

        #endregion properties
    }

    public class SettingsCollection
    {

        private Dictionary<string, Setting> _settingDictionary; //dictionary used for extra fast access to settings
        private WorksheetWithSettings _baseWorksheetWithSettings;

        public SettingsCollection(WorksheetWithSettings baseWorksheet)
        {
            _settingDictionary = new Dictionary<string, Setting>();
            this._baseWorksheetWithSettings = baseWorksheet;
        }


        public Setting this[string index]
        {
            get
            {
                if (_settingDictionary.ContainsKey(index))
                {
                    return _settingDictionary[index];
                }
                else return null;
            }

            set
            {
                Setting inSetting = value;
                if (_settingDictionary.ContainsKey(inSetting.Name))
                {
                    _settingDictionary.Remove(inSetting.Name);
                }

                this._settingDictionary.Add(inSetting.Name, inSetting);
            }
        }

        public WorksheetWithSettings BaseWorksheetWithSettings { get => this._baseWorksheetWithSettings; }

        public void AddSetting(Setting inSetting)
        {
            if (_settingDictionary.ContainsKey(inSetting.Name))
            {
                _settingDictionary.Remove(inSetting.Name);
            }

            this._settingDictionary.Add(inSetting.Name, inSetting);
        }

        public void RemoveSetting(Setting deleteSetting)
        {
            if (_settingDictionary.ContainsKey(deleteSetting.Name))
            {
                _settingDictionary.Remove(deleteSetting.Name);
            }
        }

        
                
    }

}
