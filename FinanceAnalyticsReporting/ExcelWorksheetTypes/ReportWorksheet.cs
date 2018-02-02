using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceAnalyticsReporting
{
    public class ReportWorksheet : ExcelBase.WorksheetWithSettings
    {
        public override string ReportSettingsAnchor => "reportSettings";

        public override void SaveClassSettings()
        {
           
        }
    }
}
 