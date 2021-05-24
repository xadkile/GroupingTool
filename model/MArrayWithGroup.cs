using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class MArrayWithGroup {
        public readonly object groupFlag;
        public readonly MArray array;
        public MArrayWithGroup(object groupFlag, MArray array) {
            this.groupFlag = groupFlag;
            this.array = array;
        }

        public void pourToSheet(Worksheet sheet, int row, int col) {
            this.array.pourToSheet(sheet, row, col);
        }

        public void writeToSheet(int col) {
            string groupName = this.groupFlag.ToString();
            
            Worksheet sheet = this.getWorksheetOrCreateIfDontExist(groupName);
            //Range last = (Range)sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            // get last row count
            int nextRow = Utils.getLastRow(sheet)+1;
            this.pourToSheet(sheet, nextRow, col);
        }

        private Worksheet getWorksheetOrCreateIfDontExist(string name) {
            Workbook book = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            try {
                return (Worksheet) book.Worksheets[name];
            }catch(Exception exception) {
                Worksheet rt = (Worksheet) book.Worksheets.Add();
                rt.Name = name;
                return rt;
            }
        }
    }
}
