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

            #region Enumerations and structures
            public enum DirectionType
            { Up = 3, ToRight = 2, Down = 4, ToLeft = 1 }

            #endregion Enumerations and structures


            /*Constructors - includes two overloaded constructors. One that takes a worksheet IntPtr, 
            another that is default constructor which creates a worksheet from the activesheet. */
            #region Constructors

            //constructor requires a system.intptr to hold reference to the worksheet.
            //i.e. worksheet name etc. will change
            public Worksheet(System.IntPtr worksheetIntPtr)
            {
                System.Diagnostics.Debug.WriteLine("Worksheet Base ptr ctor called!");
                this.WorkSheetPtr = worksheetIntPtr;
            }

            //assumes that the new worksheet is the active sheet
            public Worksheet()
            {

                System.Diagnostics.Debug.WriteLine("Worksheet Base default ctor called!");
                var sheetID = XlCall.Excel(XlCall.xlSheetId);
                ExcelReference newExcelRef = (ExcelReference)sheetID;
                System.IntPtr newIntPtr = newExcelRef.SheetId;
                this.WorkSheetPtr = newIntPtr;

                //TODO clean this up so refers to other constructor

            }

            #endregion Constructors

            //PROPERTIES ----------------------------------------------------------

            //save the worksheetPtr from the worksheet
            private System.IntPtr _workSheetPtr;
            public System.IntPtr WorkSheetPtr
            {
                get { return _workSheetPtr; }
                set { _workSheetPtr = value; }
            }

            private string _worksheetName;
            public string WorksheetName
            {
                get
                {
                    ExcelReference newExcelRef = new ExcelReference(0, 0, 0, 0, this.WorkSheetPtr);
                    _worksheetName = (string)XlCall.Excel(XlCall.xlSheetNm, newExcelRef);
                    return _worksheetName;
                }
                set
                {
                    //TODO - validate the worksheet name of value coming in 
                    //e.g can't name same as another worksheet, illegal characters etc.


                    //first validate that value is valid.
                    _worksheetName = value;
                }
            }

            //METHODS ----------------------------------------------------------
            public ExcelReference ReturnNamedRangeRef(string NamedRange)
            {
                string searchNamedRange = string.Format("'{0}'!{1}", this.WorksheetName, NamedRange);
                object value = XlCall.Excel(XlCall.xlfEvaluate, searchNamedRange);

                if (value is ExcelReference)
                { return (ExcelReference)value; }
                else
                { return null; }

            }

            //overloaded function can take either a named range (as as string or an excel reference)

            public ExcelReference AnchorCellToEmptySpace(string WorksheetNamedRange, DirectionType directionToLookFor)
            {
                ExcelReference namedExcelRef = this.ReturnNamedRangeRef(WorksheetNamedRange);
                return AnchorCellToEmptySpace(namedExcelRef, directionToLookFor);

            }


            public ExcelReference AnchorCellToEmptySpace(ExcelReference anchorExcelRef, DirectionType directionLookFor)
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
                    case DirectionType.Down:
                        end = new ExcelReference(anchorExcelRef.RowFirst, selection.RowLast, anchorExcelRef.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);
                        break;

                    case DirectionType.ToLeft:
                        end = new ExcelReference(anchorExcelRef.RowFirst, anchorExcelRef.RowLast, selection.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);
                        break;

                    case DirectionType.ToRight:
                        end = new ExcelReference(anchorExcelRef.RowFirst, anchorExcelRef.RowLast, anchorExcelRef.ColumnFirst, selection.ColumnLast, this.WorkSheetPtr);
                        break;

                    case DirectionType.Up:
                        end = new ExcelReference(selection.RowFirst, anchorExcelRef.RowLast, anchorExcelRef.ColumnFirst, anchorExcelRef.ColumnLast, this.WorkSheetPtr);

                        break;

                }

                return end;

            }

        }

}

