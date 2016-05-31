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
        public void Generate(string strFileName = "AXMasterSheet", int intSampleRowLines = 20, int intStartRowInput = 4, int intStartColumnInput = 2)
        {
            //開始セル
            int intStartRow = intStartRowInput;
            int intStartColumn = intStartColumnInput;

            //項目数
            int intColumnNum = 10;

            XLColor xlcBlue = XLColor.FromArgb(221, 235, 247);

            XLWorkbook.DefaultStyle.Font.FontName = "Meiryo UI";
            string strXLFileName = strFileName + ".xlsx";

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("List");

            //Headerのセル範囲を指定
            var headrange = ws.Range(intStartRow, intStartColumn, intStartRow, intStartColumn + intColumnNum - 1);

            ws.Cell(intStartRow, intStartColumn).Value = "No.";
            ws.Cell(intStartRow, intStartColumn + 1).Value = "タブ1";
            ws.Cell(intStartRow, intStartColumn + 2).Value = "タブ2";
            ws.Cell(intStartRow, intStartColumn + 3).Value = "タブ3";
            ws.Cell(intStartRow, intStartColumn + 4).Value = "タブ4";
            ws.Cell(intStartRow, intStartColumn + 5).Value = "項目";
            ws.Cell(intStartRow, intStartColumn + 6).Value = "使用/未使用";
            ws.Cell(intStartRow, intStartColumn + 7).Value = "桁数";
            ws.Cell(intStartRow, intStartColumn + 8).Value = "AX標準ヘルプ";
            ws.Cell(intStartRow, intStartColumn + 9).Value = "備考";

            //[使用/未使用]項目の選択肢を用意する
            ws.Cell(intStartRow, intStartColumn + intColumnNum + 1).Value = "使用";
            ws.Cell(intStartRow + 1, intStartColumn + intColumnNum + 1).Value = "未使用";
            ws.Range(intStartRow, intStartColumn + intColumnNum + 1, intStartRow + 1, intStartColumn + intColumnNum + 1).Style.Font.FontColor = XLColor.White;

            if (intStartColumn > 1)
            {
                for (int i = 1; i < intStartColumn; i++)
                {
                    ws.Column(i).Width = 3;
                }
            }

            ws.Column(intStartColumn).Width = 3;
            ws.Column(intStartColumn + 1).Width = 12;
            ws.Column(intStartColumn + 2).Width = 12;
            ws.Column(intStartColumn + 3).Width = 12;
            ws.Column(intStartColumn + 4).Width = 12;
            ws.Column(intStartColumn + 5).Width = 36;
            ws.Column(intStartColumn + 6).Width = 10;
            ws.Column(intStartColumn + 7).Width = 5;
            ws.Column(intStartColumn + 8).Width = 60;
            ws.Column(intStartColumn + 9).Width = 24;

            headrange.Style.Fill.BackgroundColor = xlcBlue;
            headrange.Style
                .Border.SetTopBorder(XLBorderStyleValues.Thin)
                .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                .Border.SetRightBorder(XLBorderStyleValues.Thin)
                .Border.SetBottomBorder(XLBorderStyleValues.Thin);


            //行番号を用意する
            for (int i = intStartRow + 1; i <= intSampleRowLines + intStartRow; ++i)
            {
                ws.Cell(i, intStartColumn).FormulaA1 = "ROW()-" + intStartRow.ToString();
            }


            for (int i = intStartRow + 1; i <= intSampleRowLines + intStartRow; ++i)
            {
                //一番左の枠を描画
                ws.Cell(i, intStartColumn).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

                for (int j = intStartColumn; j <= intStartColumn + intColumnNum - 1; ++j)
                {
                    ws.Cell(i, j).Style
                        .Border.SetRightBorder(XLBorderStyleValues.Thin)
                        .Border.SetBottomBorder(XLBorderStyleValues.Hair);

                    //選択肢から選ぶように設定
                    if (j == 7)
                    {
                        ws.Cell(i, j).DataValidation.List(ws.Range(intStartRow, intStartColumn + intColumnNum + 1, intStartRow + 1, intStartColumn + intColumnNum + 1));
                    }
                }
            }

            ws.Range(intSampleRowLines + intStartRow, intStartColumn, intSampleRowLines + intStartRow, intStartColumn + intColumnNum - 1).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            try
            {
                wb.SaveAs(strXLFileName);
                Console.WriteLine("AX Master Sheet was created");
            }
            catch
            {
                Console.WriteLine("There is any error while saving");
                Console.ReadLine();
            }

        }

    }
}
