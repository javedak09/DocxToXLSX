using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocxToXLSX
{
    public class FomulaCell : Cell
    {
        public FomulaCell(string header, string text, int index)
        {
            this.CellFormula = new CellFormula { CalculateCell = true, Text = text };
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.StyleIndex = 2;

        }


    }
}
