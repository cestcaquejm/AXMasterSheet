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
            Generate("AXMasterSheet");
        }

        static private void Generate(string strSheetName)
        {
            XLWorkbook.DefaultStyle.Font.FontName = "Meiryo UI";
            string strXLFileName = strSheetName + ".xlsx";
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("List");

            ws.Cell(1, 1).Value = "No.";
            ws.Cell(1, 2).Value = "タブ1";
            ws.Cell(1, 3).Value = "タブ2";
            ws.Cell(1, 4).Value = "タブ3";
            ws.Cell(1, 5).Value = "タブ4";
            ws.Cell(1, 6).Value = "項目";
            ws.Cell(1, 7).Value = "使用/未使用";
            ws.Cell(1, 8).Value = "桁数";
            ws.Cell(1, 9).Value = "概要";
            ws.Cell(1, 10).Value = "備考";

            ws.Range(1, 1, 1, 10).Style.Fill.BackgroundColor = XLColor.FromArgb(180,198,231);

            for (int i = 2; i < 22; ++i)
            {
                ws.Cell(i, 1).FormulaA1 = "ROW()-1";
            }

            ws.Columns().AdjustToContents();
            //ws.Range(1,1,21,10).Style.Font.FontName="Meiryo UI";

            try
            {
                wb.SaveAs(strXLFileName);
                Console.WriteLine("AX Master Sheet was created");
            }
            catch
            {
                Console.WriteLine("There is any error while saving");
            }

        }
    }
}
