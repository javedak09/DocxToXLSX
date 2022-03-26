using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocxToXLSX
{
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, string text, int index) : base(header, text, index)
        {
            this.StyleIndex = 2;
        }

    }
}
