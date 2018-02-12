using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBase;
using System.Diagnostics;
using System.Collections.ObjectModel;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using Xceed.Wpf.Toolkit.PropertyGrid.Editors;
using System.ComponentModel;
using System.Windows;
using Xceed.Wpf.Toolkit.PropertyGrid;

namespace FinanceAnalyticsReporting.ExcelWorksheetTypes
{
    [ExcelBase.Worksheet.WorksheetDerivedTypeIdentifier("ReportWorksheet")]
    public class ReportWorksheet : WorksheetWithSettings
    {
        public override void SettingsSavedToClass()
        {
            //ok we have 4 lists to work with, do we do anything here?

        }

        public ReportWorksheet() : base()
        {
            //when we construct this worksheet, we want to get the settings from the worksheet
            ReadSettingsToList();

        }

        public void ReloadReportWorksheet()
        {
            //First what are the most recent settings the ones in the lists or the ones on the sheet?
            //turn off screen updating
            ExcelBase.Application.TurnScreenUpdatingOff();

            //We consider the class settings the master settings.
            this.CommitAllSettingsToSheet();

            //activate and calculate the sheet
            this.Activate().Calculate();

            //List<Tuple<SettingItem, object>> listData = ActiveFormDataProvider.ReturnDataFromNamedRanges(this.listAllSettings);


            ExcelBase.Application.TurnScreenUpdatingOn();

        }

        #region properties
        private List<AssemblyX> workbookOpenPath = new List<AssemblyX>();

        //[Editor(typeof(CollectionEditor), typeof(CollectionEditor))]
        [DisplayName("ExpectedAssemblyVersions")]
        [Description("The expected assembly versions.")]
        [Category("Mandatory")]
        [ExpandableObject] 
        public List<AssemblyX> WorkbookOpenPath
        {
            get
            {


                return workbookOpenPath;

            }
            set { workbookOpenPath = value; }
        }

        [ExpandableObject] 
        public class AssemblyX
        {
            public string Name { get; set; }

            public string Version { get; set; }

            public override string ToString()
            {
                return $"{Name} {Version}";
            }
        }



        #endregion properties

    }

    
}
