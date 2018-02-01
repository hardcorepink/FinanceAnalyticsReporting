using System;
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
        private string _name;

        #endregion fields

        #region constructors

        public NamedRange(Worksheet worksheet, string name)
        {
            //we are building a worksheet scoped named range
            _worksheet = worksheet;
            _name = name;
        }

        public NamedRange(Workbook workbook, string name)
        {
            _workbook = workbook;
            _name = name;
        }

        #endregion constructors

        public string NameLocal { get => _name; }

        public bool IsLocalScope { get => _worksheet != null; }

        public bool IsGlobalScope { get => _workbook != null; }
            
        public string RefersTo { get => (string)XlCall.Excel(XlCall.xlfGetName, this.NameRef, Type.Missing); }

        public string NameRef
        {
            get
            {
                return (_workbook != null) ? _workbook.Name + "!" + _name :
              _worksheet.FullWorksheetName.Contains(" ") ?
              string.Format("'{0}'!{1}", _worksheet.FullWorksheetName, _name) :
              string.Format("{0}!{1}", _worksheet.FullWorksheetName, _name);
            }
        }


    }
}
