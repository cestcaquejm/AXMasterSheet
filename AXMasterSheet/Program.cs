using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;

namespace AXMasterSheet
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateSheet.Generate("AXMasterSheet");
        }
    }
}
