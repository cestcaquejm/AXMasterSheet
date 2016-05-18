using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace AXMasterSheet
{
    class GenerateSheet
    {
        static public void Generate(string strFileName = "AXMasterSheet", int intSampleRowLines = 20)
        {
            XLWorkbook.DefaultStyle.Font.FontName = "Meiryo UI";
            string strXLFileName = strFileName + ".xlsx";
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("List");

            //Headerのセル範囲を指定
            var headrange = ws.Range(1, 1, 1, 10);

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

            ws.Cell(1, 12).Value = "使用";
            ws.Cell(2, 12).Value = "未使用";
            ws.Range(1, 12, 2, 12).Style.Font.FontColor = XLColor.White;

            ws.Column(1).Width = 3;
            ws.Column(2).Width = 12;
            ws.Column(3).Width = 12;
            ws.Column(4).Width = 12;
            ws.Column(5).Width = 12;
            ws.Column(6).Width = 36;
            ws.Column(7).Width = 10;
            ws.Column(8).Width = 5;
            ws.Column(9).Width = 60;
            ws.Column(10).Width = 24;

            headrange.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 198, 231);
            headrange.Style
                .Border.SetTopBorder(XLBorderStyleValues.Thin)
                .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                .Border.SetRightBorder(XLBorderStyleValues.Thin)
                .Border.SetBottomBorder(XLBorderStyleValues.Thin);


            for (int i = 2; i < intSampleRowLines + 2; ++i)
            {
                ws.Cell(i, 1).FormulaA1 = "ROW()-1";
            }

            for (int i = 2; i < intSampleRowLines + 2; ++i)
            {
                ws.Cell(i, 1).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

                for (int j = 1; j < 11; ++j)
                {
                    ws.Cell(i, j).Style
                        .Border.SetRightBorder(XLBorderStyleValues.Thin)
                        .Border.SetBottomBorder(XLBorderStyleValues.Hair);

                    if (j == 7)
                    {
                        ws.Cell(i, j).DataValidation.List(ws.Range(1, 12, 2, 12));
                    }
                }
            }

            ws.Range(intSampleRowLines + 1, 1, intSampleRowLines + 1, 10).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

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
