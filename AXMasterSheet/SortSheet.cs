using System;
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
        static public int Sort(string strFileName)
        {
            Console.WriteLine("Read .xlsx file.");
            var wb = new XLWorkbook();
            try
            {
                wb = new XLWorkbook(strFileName);
            }
            catch
            {
                Console.WriteLine("Reading xlsx file was failed.");
                return 1;
            }

            const string strListSheetName = "List";
            const string strMasterSheetName = "Master";

            var wsList = wb.Worksheet(strListSheetName);
           
            try
            {
                wb.Worksheet(strMasterSheetName).Delete();
                Console.WriteLine("Master sheet was deleted.");
            }
            catch
            {
            }

            var wsMaster = wb.Worksheets.Add(strMasterSheetName);
            Console.WriteLine("New master sheet was added.");

            int intListMaxRow = wsList.RowsUsed().Count();

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

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
                dr["Use"]  = wsList.Cell(i, 7).Value;

                dt.Rows.Add(dr);
            }

            for (int i = 1; i <= dt.Rows.Count; i++)
            {
                wsMaster.Cell(i, 1).Value = dt.Rows[i - 1][0];
                wsMaster.Cell(i, 2).Value = dt.Rows[i - 1][1];
                wsMaster.Cell(i, 3).Value = dt.Rows[i - 1][2];
                wsMaster.Cell(i, 4).Value = dt.Rows[i - 1][3];
                wsMaster.Cell(i, 5).Value = dt.Rows[i - 1][4];
                wsMaster.Cell(i, 6).Value = dt.Rows[i - 1][5];
                wsMaster.Cell(i, 7).Value = dt.Rows[i - 1][6];
            }

            wb.SaveAs(strFileName);

            return 0;
        }

        static private void CopyCell(string strColumnValue,ref string strCopyValue, string strCheckColumnValue, string strCheckCopyValue)
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

    }
}
