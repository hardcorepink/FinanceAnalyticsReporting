using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBase;

namespace FinanceAnalyticsReporting
{
    [ExcelBase.Worksheet.WorksheetDerivedTypeIdentifier("ReportWorksheet")]
    public class ReportWorksheet : ExcelBase.WorksheetWithSettings
    {
        public override string ReportSettingsAnchor => "reportSettings";

        public override void SaveClassSettings()
        {

        }

        public ReportWorksheet() : base()
        {

        }
    }
}
