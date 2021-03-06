﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;

namespace AXMasterSheet
{
    class SortSheet
    {
        static public int Sort(string strFilePath, string strMode = "A")
        {
            XLColor xlcBlue = XLColor.FromArgb(221, 235, 247);
            XLColor xlcGlay = XLColor.FromArgb(217, 217, 217);

            Console.Write("Read .xlsx file...");

            var wb = new XLWorkbook();
            try
            {
                wb = new XLWorkbook(strFilePath);
                Console.WriteLine("Sucess!");
            }
            catch
            {
                Console.WriteLine("Failed.\nReading xlsx file was failed.");
                return -1;
            }

            const string strMasterSheetName = "Master";

            try
            {
                wb.Worksheet(strMasterSheetName).Delete();
                Console.WriteLine("Current " + strMasterSheetName + " sheet was deleted.");
            }
            catch
            {
            }

            var wsMaster = wb.Worksheets.Add(strMasterSheetName);
            Console.WriteLine("New Master sheet was added.");

            DataTable dtMaster = GetDataTabe(strFilePath, wb);
            DataTable dtUse = dtMaster;

            //使用する項目のみに絞り込む
            if (strMode == "U")
            {
                dtUse = SelectUse(dtMaster);
            }

            int intUseRow = dtUse.Rows.Count;

            //ヘッダーの生成
            for (int i = 1; i <= intUseRow; i++)
            {
                wsMaster.Cell(1, i + 1).Value = dtUse.Rows[i - 1][1];
                wsMaster.Cell(2, i + 1).Value = dtUse.Rows[i - 1][2];
                wsMaster.Cell(3, i + 1).Value = dtUse.Rows[i - 1][3];
                wsMaster.Cell(4, i + 1).Value = dtUse.Rows[i - 1][4];
                wsMaster.Cell(5, i + 1).Value = dtUse.Rows[i - 1][5];

                if (dtUse.Rows[i - 1][6].ToString() == "使用")
                {
                    wsMaster.Cell(5, i + 1).Style.Fill.BackgroundColor = xlcBlue;
                }
                else
                {
                    wsMaster.Cell(5, i + 1).Style.Fill.BackgroundColor = xlcGlay;
                }
            }

            //ヘッダーの塗りつぶし
            for (int i = 1; i <= intUseRow + 1; i++)
            {
                for (int j = 1; j <= 4; j++)
                {
                    wsMaster.Cell(j, i).Style.Fill.BackgroundColor = xlcBlue;
                }
            }

            //ヘッダーの内部の枠線追加
            WriteColumnLines(wsMaster, 1, intUseRow);
            WriteColumnLines(wsMaster, 2, intUseRow);
            WriteColumnLines(wsMaster, 3, intUseRow);
            WriteColumnLines(wsMaster, 4, intUseRow);

            wsMaster.Range(5,1,5,intUseRow+1).Style
                .Border.SetTopBorder(XLBorderStyleValues.Thin)
                .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                .Border.SetRightBorder(XLBorderStyleValues.Thin)
                .Border.SetBottomBorder(XLBorderStyleValues.Thin);

            //ヘッダーの外周の枠線追加
            wsMaster.Range(1, 1, 1, intUseRow + 1).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
            wsMaster.Range(1, 1, 4, 1).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
            wsMaster.Range(1, intUseRow + 2, 4, intUseRow + 2).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
            wsMaster.Range(4, 1, 4, intUseRow + 1).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            //ヘッダーのセル結合
            MergeCells(wsMaster, intUseRow, 1);
            MergeCells(wsMaster, intUseRow, 2);
            MergeCells(wsMaster, intUseRow, 3);
            MergeCells(wsMaster, intUseRow, 4);

            

            //不要なヘッダー行の削除
            int intMinusRows = 0;

            if (dtUse.Select("Tab4 <> ''").Length == 0)
            {
                wsMaster.Row(4).Delete();
                intMinusRows += 1;
            }
            if (dtUse.Select("Tab3 <> ''").Length == 0)
            {
                wsMaster.Row(3).Delete();
                intMinusRows += 1;
            }
            if (dtUse.Select("Tab2 <> ''").Length == 0)
            {
                wsMaster.Row(2).Delete();
                intMinusRows += 1;
            }

            //行番号追加
            AddLineNum(wsMaster, intMinusRows, intUseRow);

            //列幅自動調整 - 日本語文字列だと列幅調整がうまく働かない?
            wsMaster.ColumnsUsed().AdjustToContents();

            //未使用列の非表示
            //HideLine(wsMaster,intUseRow, intMinusRows);

            wb.SaveAs(strFilePath);

            return 0;
        }

        static private DataTable GetDataTabe(string strFilePath, XLWorkbook wb)
        {
            DataTable dt = new DataTable();

            const string strListSheetName = "List";
            var wsList = wb.Worksheet(strListSheetName);

            int intListMaxRow = wsList.RowsUsed().Count();

            dt.Columns.Add("Numb"); //[0]
            dt.Columns.Add("Tab1"); //[1]
            dt.Columns.Add("Tab2"); //[2]
            dt.Columns.Add("Tab3"); //[3]
            dt.Columns.Add("Tab4"); //[4]
            dt.Columns.Add("Item"); //[5]
            dt.Columns.Add("Use");  //[6]

            string strCopyValue1 = "";
            string strCopyValue2 = "";
            string strCopyValue3 = "";
            string strCopyValue4 = "";

            for (int i = 2; i <= intListMaxRow; i++)
            {
                DataRow dr = dt.NewRow();

                dr[0] = i - 1;

                string strColumnValue1 = wsList.Cell(i, 2).Value.ToString();
                string strColumnValue2 = wsList.Cell(i, 3).Value.ToString();
                string strColumnValue3 = wsList.Cell(i, 4).Value.ToString();
                string strColumnValue4 = wsList.Cell(i, 5).Value.ToString();

                if (strColumnValue1 != "")
                {
                    strCopyValue1 = strColumnValue1;
                }

                CopyCell(strColumnValue2, ref strCopyValue2, strColumnValue1, strCopyValue1);
                CopyCell(strColumnValue3, ref strCopyValue3, strColumnValue2, strCopyValue2);
                CopyCell(strColumnValue4, ref strCopyValue4, strColumnValue3, strCopyValue3);

                dr[1] = strCopyValue1;
                dr[2] = strCopyValue2;
                dr[3] = strCopyValue3;
                dr[4] = strCopyValue4;

                dr["Item"] = wsList.Cell(i, 6).Value;
                dr["Use"] = wsList.Cell(i, 7).Value;

                dt.Rows.Add(dr);
            }

            return dt;
        }

        static private void CopyCell(string strColumnValue, ref string strCopyValue, string strCheckColumnValue, string strCheckCopyValue)
        {
            if (strCheckColumnValue == strCheckCopyValue)
            {
                strCopyValue = "";
            }

            if (strColumnValue != "")
            {
                strCopyValue = strColumnValue;
            }
        }

        static private DataTable SelectUse(DataTable dtMaster)
        {
            DataTable dtUse = dtMaster.Clone();

            DataRow[] dr = dtMaster.Select("Use='使用'");

            foreach (DataRow rows in dr)
            {
                DataRow drClone = dtUse.NewRow();

                drClone.ItemArray = rows.ItemArray;

                dtUse.Rows.Add(drClone);
            }

            return dtUse;
        }

        static private IXLWorksheet AddLineNum(IXLWorksheet ws, int intMinusRows, int intUseRow, int intLines = 50)
        {
            int intStartRow = 5 - intMinusRows;
            string strFormula = "ROW() -" + intStartRow.ToString();

            ws.Cell(intStartRow, 1).Value = "No.";
            ws.Cell(intStartRow, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(221, 235, 247);

            for (int i = 1; i <= intLines; i++)
            {
                ws.Cell(i + intStartRow, 1).FormulaA1 = strFormula;

                for (int j = 1; j <= intUseRow + 1; j++)
                {
                    ws.Cell(i + intStartRow, j).Style
                        .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                        .Border.SetRightBorder(XLBorderStyleValues.Thin)
                        .Border.SetBottomBorder(XLBorderStyleValues.Hair);
                }

                ws.Range(intStartRow + intLines, 1, intStartRow + intLines, 1 + intUseRow).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

            }
            return ws;
        }

        static private void WriteColumnLines(IXLWorksheet ws, int intEffectRow, int intMaxRow)
        {
            string strBoarderCheck = "";

            ws.Cell(intEffectRow, 2).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);

            for (int i = 2; i <= intMaxRow + 1; i++)
            {
                string strCurrentRowValue = ws.Cell(intEffectRow, i).Value.ToString();

                if (strCurrentRowValue != strBoarderCheck)
                {
                    ws.Range(intEffectRow, i, 4, i).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
                }

                if (strCurrentRowValue != "")
                {
                    ws.Cell(intEffectRow, i).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                }

                strBoarderCheck = strCurrentRowValue;
            }
        }

        static private void MergeCells(IXLWorksheet ws, int intUseRow, int intCheckRow)
        {
            int intCount = 0;

            //ヘッダーのセル結合
            for (int i = 2; i <= intUseRow + 2; i++)
            {
                if (ws.Cell(intCheckRow, i).Style.Border.LeftBorder == XLBorderStyleValues.Thin)
                {
                    ws.Range(intCheckRow, i - intCount, intCheckRow, i - 1).Merge();
                    intCount = 1;
                }
                else
                {
                    intCount += 1;
                }
            }
        }

        static private void HideLine(IXLWorksheet ws, int intUseRow, int intMinusRows)
        {
            int intStartRow = 5 - intMinusRows;

            for (int i = 1; i <= intUseRow + 1; i++)
            {
                if (ws.Cell(intStartRow,i).Style.Fill.BackgroundColor == XLColor.FromArgb(217, 217, 217))
                {
                    ws.Column(i).Hide();
                }
            }

        }
    }
}
