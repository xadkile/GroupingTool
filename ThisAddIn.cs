using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using Microsoft.Office.Interop.Excel;
using GroupingTool.model;

namespace GroupingTool {
    public partial class ThisAddIn {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            this.Application.WorkbookBeforeSave 
                += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(zzzz);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
        }
        void zzzz(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel){
            //object[][] data; 
            //data = { { "z"} };
            //object[] array = { "A", "B" };
            object[,] array2 = { { "x1", "x2" }, { "x3", "x4" } };
            Worksheet currentSheet = (Worksheet)Wb.ActiveSheet;
            Range r = currentSheet.Range["A1:K1"];
            object[,] t = (object[,])r.Value2;
            object[][] tj = Utils.toJaggedArray(t,true);

            object[] rxxx = tj[0];



            Range last = (Range) currentSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string lastAddress = last.Address;
            

            ((Range)((Worksheet) Wb.ActiveSheet).Range["E1:G2"]).Value2 = array2;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
