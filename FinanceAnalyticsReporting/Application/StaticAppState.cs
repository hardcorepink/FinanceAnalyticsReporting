using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Diagnostics;
using System.Reflection;
using ExcelBase;

//TODO need to implement INotifyPropertyChanged - but for now working without GUI - just getting things working perfectly before GUI move....

namespace FinanceAnalyticsReporting.Application
{
    public static class StaticAppState
    {

        private static Dictionary<string, Type> dictionaryWSTypesFromStrings;
        private static ExcelBase.Worksheet _currentActiveWorkSheet;
        
        /// <summary>
        /// This is part of the startup routines. Reflects over types that derive from ExcelBase Worksheet and
        /// looks for specific identifier in cell A1 in Attributes. For example if ReportWorksheet type is in A1, then
        /// _specificWorksheetTYpe should receive an instance of RpeortWorksheet.
        /// </summary>
        public static void BuildDictionaryOfWorksheetTypes()
        {
            dictionaryWSTypesFromStrings = new Dictionary<string, Type>();
            //get the current assembly
            Assembly thisAssembly = Assembly.GetExecutingAssembly();

            Type[] ListTypesInThisAssembly = thisAssembly.GetTypes();

            foreach (Type t in ListTypesInThisAssembly)
            {
                if (t.IsSubclassOf(typeof(ExcelBase.Worksheet)))
                {
                    object[] customAtts = t.GetCustomAttributes(false);
                    foreach (ExcelBase.Worksheet.WorksheetDerivedTypeIdentifierAttribute a1Ref in customAtts)
                    {
                        dictionaryWSTypesFromStrings.Add(a1Ref.ClassIdentifierString, t);
                    }

                }
            }
        }

        [ExcelCommand()]
        public static void WorksheetChangedApp()
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            //worksheet has changed so clear the current worksheet variables
            _currentActiveWorkSheet = null;

            try
            {
                //set the _currentActiveWorksheet (base class type)
                _currentActiveWorkSheet = new ExcelBase.Worksheet();

                //get the value in a1 of the currentActiveSheet
                ExcelReference newExcelRef = new ExcelReference(0, 0, 0, 0, _currentActiveWorkSheet.WorkSheetPtr);
                object valueA1 = newExcelRef.GetValue();
                string valueA1String = valueA1.ToString();

                //here try and create a more specific worksheet type
                if (dictionaryWSTypesFromStrings.ContainsKey(valueA1String))
                {
                    Type specificWSType = dictionaryWSTypesFromStrings[valueA1String];
                    _currentActiveWorkSheet = (ExcelBase.Worksheet)Activator.CreateInstance(specificWSType);
                }

                Debug.WriteLine($"Type of currentworksheet is {_currentActiveWorkSheet.GetType().Name}");
            }
            catch
            {
                Debug.WriteLine($"Selected sheet is not of worksheet type");
            }

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            Debug.WriteLine($"Done in {elapsedMs} ms");

        }

    }
}
