using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class NamedRange
    {
        #region fields

        private Worksheet _worksheet;
        private Workbook _workbook;
        private string _shortName;

        #endregion fields

        #region constructors

        public NamedRange(Worksheet worksheet, string shortName)
        {
            //we are getting a worksheet scoped name range - this will get a hadle to the name
            //if it doesn't exist will return error
            this._worksheet = worksheet;
            this._shortName = shortName;
        }

        public NamedRange(Workbook workbook, string shortName)
        {
            this._workbook = workbook;
            this._shortName = shortName;
        }

        #endregion constructors

        #region properties

        public string ShortName
        {
            get { return _shortName; }
        }
        #endregion properties

        #region methods

        public void Delete()
        {
            //We cannot use fullName here as Delete.Name does not recongnise names in the format '[Book1]'!NamedRange
            string nameDeleteFormat;

            //TODO - ensure that the worksheet/workbook that this name is referring to is the active one.

            if(this._workbook == null)
            {
                //worksheet scoped
                nameDeleteFormat = $"'{this._worksheet.ShortWorksheetName}'!{this.ShortName}";
            }
            else
            {
                //workbook scoped name
                nameDeleteFormat = this.ShortName;

                //todo need to test if there is a 'competing' worksheet name in here. If so add sheet, delete name then delete new sheet
            }

            XlCall.Excel(XlCall.xlcDeleteName, nameDeleteFormat);
        }

        public string FullName
        {
            get
            {
                if (_workbook == null)
                {
                    //worksheet scoped but full name will still include the workbook '[Workbook.xlsx]Work Sheet'!WorksheetScopedNamedRange
                    return $"'[{this._worksheet.ParentWorkbook.Name}]{this._worksheet.ShortWorksheetName}'!{this.ShortName}";
                }
                else
                {
                    //workbook scoped in format ='Workbook.xlsx'!WorkbookScopedNamedRange
                    return $"'{this._workbook.Name}'!{this.ShortName}";
                }
            }
        }

        #endregion methods

    }
    public class NamedRangeCollection : IEnumerable<NamedRange>
    {

        private Worksheet _worksheet;
        private Workbook _workbook;


        public NamedRange this[string shortName]
        {
            //remember Names function only returns a short name, so this must be passed a short name
            get
            {
                foreach (NamedRange nr in this)
                {
                    if (nr.ShortName == shortName) return nr;
                }
                return null;
            }
        }



        public NamedRangeCollection(Workbook workbookOfNames)
        {
            this._workbook = workbookOfNames;
        }

        public NamedRangeCollection(Worksheet worksheetOfNames)
        {
            this._worksheet = worksheetOfNames;
        }

        public IEnumerator<NamedRange> GetEnumerator()
        {
            //work out the document name - if we are looking for worksheet names then doc is '[book1]sheet1', otherwsie if workbook names doc is 'book1'
            string documentName;
            object[,] arrayOfNames;

            if (_worksheet == null) { documentName = _workbook.Name; }
            else { documentName = $"[{_worksheet.ParentWorkbook.Name}]{_worksheet.ShortWorksheetName}"; }


            try
            {
                arrayOfNames = (object[,])XlCall.Excel(XlCall.xlfNames, documentName);
            }
            catch
            {
                yield break;
            }


            int numNames = arrayOfNames.GetLength(1);


            for (int i = 0; i < numNames; i++)
            {
                //for workbook defined names we are only returning workbook scoped names
                //for worksheet defined names we are only returing worksheet scoped names (_worksheet)
                if (_worksheet != null)
                {
                    bool trueIfworksheetScoped = false;

                    try
                    {
                        string worksheetScopedName = $"'[{this._worksheet.ParentWorkbook.Name}]{this._worksheet.ShortWorksheetName}'!{(string)arrayOfNames[0, i]}";
                        trueIfworksheetScoped = (bool)XlCall.Excel(XlCall.xlfGetName, worksheetScopedName, 2);
                    }
                    catch
                    {
                        continue;
                    }
                    if (trueIfworksheetScoped) { yield return new NamedRange(this._worksheet, (string)arrayOfNames[0, i]); }

                    //we have an array of worksheet scoped names - as we passed the sheet name

                }
                else
                {
                    bool falseIfWorkbookScoped = true;
                    try
                    {
                        //we only want to return workbook scoped names
                        string workbookScopedName = $"'{this._workbook.Name}'!{(string)arrayOfNames[0, i]}";
                        falseIfWorkbookScoped = (bool)XlCall.Excel(XlCall.xlfGetName, workbookScopedName, 2);
                    }
                    catch
                    {
                        continue;
                    }
                    if (falseIfWorkbookScoped == false) { yield return new NamedRange(this._workbook, (string)arrayOfNames[0, i]); }
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (NamedRange nr in this)
            {
                sb.Append(nr.ShortName);
                sb.Append(Environment.NewLine);
            }

            return sb.ToString();

        }

        public NamedRange Add(string nameText, string refersTo, bool hidden = true)
        {
            try
            {
                //TODO need to activate the workbook passed here, as defineName assumes active workbook
                bool local = false;

                if (_worksheet == null) { local = false; }
                else { local = true; }

                XlCall.Excel(XlCall.xlcDefineName, nameText, refersTo, Type.Missing, Type.Missing, hidden, Type.Missing, local);

                if (_worksheet == null) { return new NamedRange(this._workbook, nameText); } //workbook scoped
                else { return new NamedRange(this._worksheet, nameText); }

            }
            catch
            {
                return null;
            }
        }


        //TODO - add named range indexer for worksheet and workbook


    }
}
