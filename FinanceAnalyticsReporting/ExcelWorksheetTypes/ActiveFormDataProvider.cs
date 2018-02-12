using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelBase;
using System.IO;
using ExcelDna.Integration;

namespace FinanceAnalyticsReporting.ExcelWorksheetTypes
{
    public static class ActiveFormDataProvider
    {

        //what do we return from here???
        //public static List<Tuple<SettingItem, object>> ReturnDataFromNamedRanges(List<SettingItem> allSettinsIn)
        //{

        //    Workbook defaultWorkbookOpen = null;

        //    //get data workbook
        //    string dataWorkbookPath = allSettinsIn.FirstOrDefault(si =>
        //        String.Equals(si.SettingType.ToString(), "ClassSetting", StringComparison.OrdinalIgnoreCase) &&
        //        String.Equals(si.SettingName.ToString(), "dataWorkbook", StringComparison.OrdinalIgnoreCase)).SettingValue.ToString();

        //   List<string> namedDataReferences = allSettinsIn.Where(si =>
        //        String.Equals(si.SettingType.ToString(), "namedData", StringComparison.OrdinalIgnoreCase)).Select(si => si.SettingValue.ToString()).ToList();


        //    List<ExcelReference> = namedDataReferences.ForEach(namedDataString => XlCall.Excel(XlCall.xlfEvaluate, namedDataString)
                                   
            

        //    //now reload every sheet
        //    List<ExcelReference> uniqueExcelReferences = listExcelReferences.Distinct().ToList


        //    return null;
        //}


    }

    public class NamedRangeRef
    {
        //public string WorkbookFullFIlePathString { get; set; }
        //public string WorkbookNameString { get; set; }
        //public string WorksheetName { get; set; }
        //public string FullSheetReference { get; set; }
        //public string RangeAddressString { get; set; }


        //public NamedRangeRef(string fullNamedRangeReference, Workbook defaultWorkbook, SettingItem settingItemRef)
        //{
        //    //we have a string for the fullNamedRangeReference...
        //    //this is either one of 2 formats
        //    //With full workbook path or without
        //    int indexerOfLastEplanationPoint = fullNamedRangeReference.LastIndexOf("!");
        //    RangeAddressString = fullNamedRangeReference.Substring(indexerOfLastEplanationPoint, fullNamedRangeReference.Length - indexerOfLastEplanationPoint);

        //    //
        //    string remainingString = fullNamedRangeReference.Substring(0, indexerOfLastEplanationPoint - 1);





        //}


    }


}


