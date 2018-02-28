using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class Worksheet
    {

        #region fields

        private System.IntPtr _workSheetPtr;
        private ExcelDna.Integration.ExcelReference _baseExcelReference;
        private RangeClass _range;

        #endregion fields

        #region Constructors

        /// <summary>
        /// Constructs a Worksheet from the currently active sheet.
        /// </summary>
        public Worksheet() : this(((ExcelReference)XlCall.Excel(XlCall.xlSheetId)).SheetId)
        {
        }

        /// <summary>
        /// Constructs a Worksheet from a full worksheet reference. For example [Book1]SheetName
        /// </summary>
        public Worksheet(string FullWorksheetReference) : this(((ExcelReference)XlCall.Excel(XlCall.xlSheetId, FullWorksheetReference)).SheetId)
        {
        }

        public Worksheet(IntPtr sheetID)
        {
            this._baseExcelReference = new ExcelReference(0, 0, 0, 0, sheetID);
            this._workSheetPtr = this._baseExcelReference.SheetId;
            this._range = new RangeClass(this);
        }


        #endregion Constructors

        #region AttributeClass

        [System.AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
        public sealed class WorksheetDerivedTypeIdentifierAttribute : Attribute
        {
            readonly string _a1Identifier;

            // This is a positional argument
            public WorksheetDerivedTypeIdentifierAttribute(string ClassIdentifierString)
            {
                this._a1Identifier = ClassIdentifierString;
            }

            public string ClassIdentifierString
            {
                get { return _a1Identifier; }
            }


        }


        #endregion AttributeClasses

        #region Properties

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
                Workbook newWB = new Workbook(workbookName);
                return newWB;

            }
        }

        public RangeClass Range
        {
            get { return this._range; }
        }


        public NamedRangeCollection Names
        {
            get
            {
                return new NamedRangeCollection(this);
            }
        }


        #endregion Properties

        #region Methods

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

        /// <summary>
        /// Activates the worksheet instance in excel, returns the Worksheet reference 
        /// so can chaing together actions e.g. Worksheet.Activate().Calculate()
        /// </summary>
        public Worksheet Activate()
        {
            XlCall.Excel(XlCall.xlcWorkbookActivate, this.FullWorksheetName);
            return this;
        }

        /// <summary>
        /// Calculates the active worksheet. Make sure that this worksheet is active before calculating using
        /// Worksheet.Activate()
        /// </summary>
        public Worksheet Calculate()
        {
            XlCall.Excel(XlCall.xlcCalculateDocument);
            return this;
        }

        #endregion Methods

        #region classes

        public class WorksheetsCollection : IEnumerable<Worksheet>
        {
            private string _workbookName;

            public WorksheetsCollection(string WorkbookName)
            {
                this._workbookName = WorkbookName;
            }

            public Worksheet this[string worksheetName]
            {
                get
                {
                    try
                    {   //Get Workbook given 1 returns array of worksheet names in format [BookName]SheetName                     
                        object[,] ArrayFullWorksheetNames = (object[,])(XlCall.Excel(XlCall.xlfGetWorkbook, 1, this._workbookName));
                        long numberSheets = ArrayFullWorksheetNames.GetLongLength(1);
                        for (int i = 0; i < numberSheets; i++)
                        {
                            //get just the sheet name
                            string sheetName = Worksheet.WorksheetNameFromFullReference((string)ArrayFullWorksheetNames[0, i]);
                            //compare to what we provided
                            if (String.Equals(worksheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    //construct a new worksheet type from the full worksheet name
                                    return new Worksheet((string)ArrayFullWorksheetNames[0, i]);
                                }
                                catch
                                {
                                    return null;
                                }
                            }
                        }

                        return null;

                    }
                    catch { return null; }
                    //this will loop through workbook worksheets. 

                }
            }

            public IEnumerator<Worksheet> GetEnumerator()
            {
                object[,] ArrayFullWorksheetNames = (object[,])(XlCall.Excel(XlCall.xlfGetWorkbook, 1, this._workbookName));
                long numberSheets = ArrayFullWorksheetNames.GetLongLength(1);
                for (int i = 0; i < numberSheets; i++)
                {
                    //always need to try constructing a worksheet, just in case it can't work (chart sheet for excample)
                    Worksheet WorksheetToReturn;
                    try
                    {
                        WorksheetToReturn = new Worksheet((string)ArrayFullWorksheetNames[0, i]);
                    }
                    catch
                    {
                        continue;
                    }
                    yield return WorksheetToReturn;

                }
            }
            IEnumerator IEnumerable.GetEnumerator()
            {
                return this.GetEnumerator();
            }

            public override string ToString()
            {
                StringBuilder sb = new StringBuilder();
                foreach (Worksheet ws in this)
                {
                    sb.Append(ws.FullWorksheetName);
                    sb.Append(Environment.NewLine);
                }

                return sb.ToString();

            }
        }

        #endregion classes
    }
}

