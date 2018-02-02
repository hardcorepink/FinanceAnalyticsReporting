using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Diagnostics;

//TODO need to implement INotifyPropertyChanged - but for now working without GUI - just getting things working perfectly before GUI move....

namespace FinanceAnalyticsReporting.Application
{
    public static class StaticAppState
    {
        private static ExcelBase.Worksheet _currentActiveSheet;

        public static ExcelBase.Worksheet CurrentActiveSheet
        {
            get { return _currentActiveSheet; }
            set { _currentActiveSheet = value; }
        }

        [ExcelCommand()]
        public static void WorksheetChangedApp()
        {
            _currentActiveSheet = new ExcelBase.Worksheet();
            Debug.WriteLine($"Worksheet Changed to : {_currentActiveSheet.FullWorksheetName}");

        }




    }
}
