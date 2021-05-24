using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class MainLogic {
        public static void run(InputModel input) {
            Range c1 = (Range)input.dataSheet.Cells[input.fromRow, input.fromCol];
            Range c2 = (Range)input.dataSheet.Cells[input.toRow, input.toCol];
            Range inputRange =(Range) input.dataSheet.Range[c1, c2];

            MArray array = MArray.createFromRange(inputRange, input.groupByIndex, input.sortIndices);
            foreach(MArrayWithGroup mg in array.splitByGroupFlag()) {
                mg.writeToSheet(input.fromCol);
            }
        }
    }
}
