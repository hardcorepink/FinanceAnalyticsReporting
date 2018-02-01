using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelBase
{
    public class Application
    {
        //Workbooks collection 
        //use Documents on page 104 to get all workbooks

        public static List<String> ListWorkbookNames()
        {
            //see page 104 Documents returns a horizontall array in text form he names of the open workbooks
            object arrayOfWorkbooks = XlCall.Excel(XlCall.xlfDocuments, 3);
            object[,] objectArrayOfWorkbooks = (object[,])arrayOfWorkbooks;
            List<string> listWorkbooks = new List<string>();

            for(int i = 0; i < objectArrayOfWorkbooks.GetLength(1); i++)
            {
                listWorkbooks.Add((string)objectArrayOfWorkbooks[0, i]);
            }

            return listWorkbooks;
        }
    }

    //screen updating
    //use Echo on page 105 to turn on and off screen updating

}

