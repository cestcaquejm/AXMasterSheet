using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace AXMasterSheet
{
    class Program
    {
        static int Main(string[] args)
        {
            GenerateSheet gs = new GenerateSheet();
            gs.Generate("AXMasterSheet");
            int intReturn = SortSheet.Sort("pythontest.xlsx", "A");
            return intReturn;
        }
    }
}
