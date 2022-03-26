using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocxToXLSX
{
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {

            this.DataType = CellValues.Date;
            this.CellReference = header + index;
            this.StyleIndex = 1;
            this.CellValue = new CellValue { Text = dateTime.ToOADate().ToString() }; ;

        }

    }
}
