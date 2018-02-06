using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceAnalyticsReporting.ExcelWorksheetTypes
{
    [ExcelBase.Worksheet.WorksheetDerivedTypeIdentifier("ActiveFormDataSheet")]
    public class ActiveFormDataSheet : ExcelBase.WorksheetWithSettings
    {

        //this is the default constructor
        public ActiveFormDataSheet()
        {
            
        }

        

        public override void SettingsSavedToClass()
        {
            throw new NotImplementedException();
        }



    }
}
