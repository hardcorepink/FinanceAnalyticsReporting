using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class RangeClass
    {
        private Worksheet _worksheet;

        internal RangeClass(Worksheet ws)
        {
            this._worksheet = ws;
        }

        public ExcelReference this[string index]
        {
            get
            {
                ExcelReference _excelReference = null;

                try
                {
                    _excelReference = (ExcelReference)XlCall.Excel(XlCall.xlfTextref, index, true);
                }
                catch
                {
                    try
                    {
                        _excelReference = (ExcelReference)XlCall.Excel(XlCall.xlfTextref, index, false);
                    }

                    catch
                    {
                        try
                        {
                            string newIndex = @"'" + this._worksheet.ShortWorksheetName + @"'!" + index;
                            _excelReference = (ExcelReference)XlCall.Excel(XlCall.xlfTextref, newIndex, true);
                        }
                        catch
                        {
                            return null;
                        }
                    }
                }

                int[][] newIntArray = new int[_excelReference.InnerReferences.Count][];
                newIntArray[0] = new int[] { _excelReference.RowFirst, _excelReference.RowLast, _excelReference.ColumnFirst, _excelReference.ColumnLast };

                for (int i = 0; i < _excelReference.InnerReferences.Count - 1; i++)
                {
                    newIntArray[i + 1] = new int[] { _excelReference.InnerReferences[i + 1].RowFirst, _excelReference.InnerReferences[i + 1].RowLast, _excelReference.InnerReferences[i + 1].ColumnFirst, _excelReference.InnerReferences[i + 1].ColumnLast };
                }

                return new ExcelReference(newIntArray, this._worksheet.WorkSheetPtr);

            }

        }


    }

    public static class ExcelRefExtenstions
    {
        //extension methods for ExcelReference -just have the test at the moment
        public static int TestExtMethod(this ExcelReference excelRef)
        {
            return 99;
        }
    }

}
