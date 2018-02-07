using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBase;
using System.Diagnostics;

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

            List<Tuple<SettingItem, object>> listData = ActiveFormDataProvider.ReturnDataFromNamedRanges(this.listAllSettings);


            ExcelBase.Application.TurnScreenUpdatingOn();
  
        }

        #region properties



        #endregion properties

    }
}
