using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class Worksheet

    {

        #region Constructors
        
        /// <summary>
        /// Constructs a Worksheet from the currently active sheet.
        /// </summary>
        public Worksheet()
        {
            //we only care about worksheet types here. If we get an exception then we are not pointing to a worksheet

            try
            {
                object test = XlCall.Excel(XlCall.xlSheetId);
                this._baseExcelReference = (ExcelReference)XlCall.Excel(XlCall.xlSheetId);
                this._workSheetPtr = this._baseExcelReference.SheetId;
            }
            catch
            {
                throw new Exception("Could not create worksheet. Active Sheet is not worksheet type");
            }

        }

        /// <summary>
        /// Constructs a Worksheet from a full worksheet reference. For example [Book1]SheetName
        /// </summary>
        public Worksheet(string FullWorksheetReference)
        {
            try
            {
                this._baseExcelReference = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, FullWorksheetReference);
                this._workSheetPtr = this._baseExcelReference.SheetId;
            }
            catch
            {
                throw new Exception($"Could not create workbook from full worksheet reference: {FullWorksheetReference}");
            }

        }
               

        #endregion Constructors

        #region fields

        private System.IntPtr _workSheetPtr;
        private ExcelDna.Integration.ExcelReference _baseExcelReference;




        #endregion fields

        //PROPERTIES ----------------------------------------------------------

        //be careful as a sheetptr is not persistant, it may change if the sheet is moved...
        public System.IntPtr WorkSheetPtr { get => _workSheetPtr; }

        public string WorkbookName
        {
            get
            {
                //get worksheetFullName
                string worksheetFullName = this.FullWorksheetName;

                //clean the string to remove Book name e.g. [Book1]Sheet1 becomes Sheet1
                int index = worksheetFullName.IndexOf("]");

                if (index > 0)
                    return worksheetFullName.Substring(1, worksheetFullName.Length - (worksheetFullName.Length - index) - 1);
                else
                    return worksheetFullName;
            }
        }

        public string ShortWorksheetName
        {
            get
            {
                string worksheetFullName = this.FullWorksheetName;

                //clean the string to remove Book name e.g. [Book1]Sheet1 becomes Sheet1
                int index = worksheetFullName.IndexOf("]");

                if (index > 0)
                    return worksheetFullName.Substring(index + 1, worksheetFullName.Length - index - 1);
                else
                    return worksheetFullName;
            }
            set
            {
                //TODO - validate the worksheet name of value coming in 
                //e.g can't name same as another worksheet, illegal characters etc.

                //first validate that value is valid.                
            }
        }

        public string FullWorksheetName
        {
            get
            {
                //get an excel reference to pass to sheetNm Formula
                ExcelReference newExcelRef = new ExcelReference(0, 0, 0, 0, this.WorkSheetPtr);

                //get the full name of the worksheet
                return (string)XlCall.Excel(XlCall.xlSheetNm, newExcelRef);
            }
        }

        public Workbook ParentWorkbook
        {
            get
            {
                //first get the workbook name
                string workbookName = this.WorkbookName;

                Workbook newWB = new Workbook(workbookName, true);

                return newWB;
                //then get the list of 

            }
        }


        //METHODS ----------------------------------------------------------
        public ExcelReference ReturnNamedRangeRef(string NamedRange)
        {
            string searchNamedRange = string.Format("'{0}'!{1}", this.ShortWorksheetName, NamedRange);
            object value = XlCall.Excel(XlCall.xlfEvaluate, searchNamedRange);

            if (value is ExcelReference)
            { return (ExcelReference)value; }
            else
            { return null; }
        }

        //overloaded function can take either a named range (as as string or an excel reference)

        public ExcelReference AnchorCellToEmptySpace(string WorksheetNamedRange, ExcelEnums.DirectionType directionToLookFor)
        {
            ExcelReference namedExcelRef = this.ReturnNamedRangeRef(WorksheetNamedRange);
            return AnchorCellToEmptySpace(namedExcelRef, directionToLookFor);
        }

        public ExcelReference AnchorCellToEmptySpace(ExcelReference anchorExcelRef, ExcelEnums.DirectionType directionLookFor)
        {
            //based on the reference recieved will look left, right, up or down for cells until empty
            //the returns range from anchor cell to one before empty cell
            //if parameter reference is an empty cell will return from empty cell to (and including) first cell with value

            //first set end to null just incase don't get a proper ExcelReference 
            ExcelReference end = null;

            //we can only do this by range selection - need to consider locked worksheets, screen updating etc
            XlCall.Excel(XlCall.xlcSelect, anchorExcelRef);

            //the reference is now selected - select from there to last cell with value
            XlCall.Excel(XlCall.xlcSelectEnd, (int)directionLookFor);

            //get the new selected range
            var selection = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;

            //make a new reference which is from paramater range through to last non-empty cell
            switch (directionLookFor)
            {
                case ExcelEnums.DirectionType.Down:
                    end = new ExcelReference(anchorExcelRef.RowFirst, selection.RowLast, anchorExcelRef.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);
                    break;

                case ExcelEnums.DirectionType.ToLeft:
                    end = new ExcelReference(anchorExcelRef.RowFirst, anchorExcelRef.RowLast, selection.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);
                    break;

                case ExcelEnums.DirectionType.ToRight:
                    end = new ExcelReference(anchorExcelRef.RowFirst, anchorExcelRef.RowLast, anchorExcelRef.ColumnFirst, selection.ColumnLast, this.WorkSheetPtr);
                    break;

                case ExcelEnums.DirectionType.Up:
                    end = new ExcelReference(selection.RowFirst, anchorExcelRef.RowLast, anchorExcelRef.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);

                    break;

            }

            return end;

        }

        public bool IsPointerStillValid()
        {
            try
            {
                string wsName = (string)XlCall.Excel(XlCall.xlSheetNm, this._baseExcelReference);
                return true;
            }
            catch { return false; }
        }

        public static string WorksheetNameFromFullReference(string fullWorksheetName)
        {
            int index = fullWorksheetName.IndexOf("]");
            if (index > 0)
            {
                return fullWorksheetName.Substring(index + 1);
            }
            else
            {
                return fullWorksheetName;
            }
        }

    }

}

