using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class InputFacade {
        string dataSheetName;
        string labelRangeAddress;
        string groupByFlag;
        ImmutableList<string> sortFlagList;
        int fromRow;
        int toRow;

        public InputFacade(string dataSheetName,
                            string labelRangeAddress,
                            string groupByFlag,
                            ImmutableList<string> sortFlagList,
                            int fromRow, int toRow) {
            this.dataSheetName = dataSheetName;
            this.labelRangeAddress = labelRangeAddress;
            this.groupByFlag = groupByFlag;
            this.sortFlagList = sortFlagList;
            this.fromRow = fromRow;
            this.toRow = toRow;
        }

        public InputModel toModel() {
            Worksheet dataSheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets[this.dataSheetName];
            Range labelRange = (Range) dataSheet.Range[this.labelRangeAddress];
            int groupByIndex = -1;
            List<string> labelList = new List<string>();
            
            int i = 0;

            foreach (Range cell in labelRange) {
                string cellStrValue = cell.Value2.ToString();
                // groupByIndex is update only once
                if (groupByIndex == -1) {
                    if (cellStrValue.Equals(this.groupByFlag)) {
                        groupByIndex = i;
                    }
                }

                labelList.Add(cellStrValue);

                i += 1;
            }
            List<int> sortIndices = new List<int>();
            foreach(string sortFlag  in this.sortFlagList) {
                for(int x =0;x < labelList.Count; ++x) {
                    if (sortFlag.Equals(labelList[x])) {
                        sortIndices.Add(x);
                    }
                }
            }

            int fromCol = labelRange.Column;
            int toCol = labelRange.Column + labelRange.Columns.Count - 1;
            int fromRow = this.getTrueFromRow(this.fromRow,dataSheet);
            int toRow = this.getTrueToRow(this.toRow, dataSheet);

            return new InputModel(dataSheet, groupByIndex, sortIndices, labelList, fromRow,toRow, fromCol, toCol);

        }
        private int getTrueFromRow(int fromRow,Worksheet sheet) {
            int firstRow = Utils.getFirstRow(sheet);
            return Math.Max(fromRow, firstRow);
        }

        private int getTrueToRow(int toRow, Worksheet sheet) {
            int lastRow = Utils.getLastRow(sheet);
            return Math.Min(lastRow, toRow);
        }
    }
}
