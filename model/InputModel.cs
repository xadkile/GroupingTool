using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class InputModel {
        public Worksheet dataSheet;
        public int groupByIndex;
        public List<int> sortIndices;
        public List<string> labelList;
        public int fromRow;
        public int toRow;
        public int fromCol;
        public int toCol;

        public InputModel(Worksheet dataSheet,
                        int groupByIndex,
                        List<int> sortIndices,
                        List<string> labelList,
                        int fromRow, int toRow,
                        int fromCol, int toCol) {
            this.dataSheet = dataSheet;
            this.groupByIndex = groupByIndex;
            this.sortIndices = sortIndices;
            this.labelList = labelList;
            this.fromRow = fromRow;
            this.toRow = toRow;
            this.fromCol = fromCol;
            this.toCol = toCol;
        }
    }
}
