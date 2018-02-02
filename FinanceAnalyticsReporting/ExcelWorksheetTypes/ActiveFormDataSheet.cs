using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceAnalyticsReporting.ExcelWorksheetTypes
{
    public class ActiveFormDataSheet : ExcelBase.WorksheetWithSettings
    {
        public override string ReportSettingsAnchor => throw new NotImplementedException();

        public override void SaveClassSettings()
        {
            throw new NotImplementedException();
        }
    }
}
