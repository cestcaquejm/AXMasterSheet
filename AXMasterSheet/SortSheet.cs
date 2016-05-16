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
        static public void Sort(string strFileName)
        {
            var wb = new XLWorkbook(strFileName);
            var wsList = wb.Worksheet("List");
           
            try
            {
                wb.Worksheet("Master").Delete();

            }
            catch
            {
            }

            var wsMaster = wb.Worksheets.Add("Master");


            int intListMaxRow = wsList.RowsUsed().Count();

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            dt.Columns.Add("Numb");
            dt.Columns.Add("Tab1");
            dt.Columns.Add("Tab2");
            dt.Columns.Add("Tab3");
            dt.Columns.Add("Tab4");
            dt.Columns.Add("Item");
            dt.Columns.Add("Use");

            string strCheckTab1 = "";
            string strCheckTab2 = "";
            string strCheckTab3 = "";
            string strCheckTab4 = "";
            string strCheckItem = "";
            string strCheckUse = "";

            for (int i = 2; i <= intListMaxRow; i++)
            {
                DataRow dr = dt.NewRow();

                dr["Numb"] = i - 1;

                if (wsList.Cell(i,2).Value.ToString() == "")
                {
                    dr["Tab1"] = strCheckTab1;                    
                }
                else
                { 
                    dr["Tab1"] = wsList.Cell(i, 2).Value.ToString();
                    strCheckTab1 = wsList.Cell(i, 2).Value.ToString();
                }

                dr["Tab2"] = wsList.Cell(i, 3).Value;
                dr["Tab3"] = wsList.Cell(i, 4).Value;
                dr["Tab4"] = wsList.Cell(i, 5).Value;
                dr["Item"] = wsList.Cell(i, 6).Value;
                dr["Use"] = wsList.Cell(i, 7).Value;

                dt.Rows.Add(dr);


            }

            for (int i = 1; i <= dt.Rows.Count; i++)
            {
                wsMaster.Cell(i, 1).Value = dt.Rows[i - 1]["Tab1"];
                wsMaster.Cell(i, 2).Value = dt.Rows[i - 1]["Tab2"];

            }

            wb.SaveAs(strFileName);
        }
    }
}
