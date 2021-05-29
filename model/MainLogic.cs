using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class MainLogic {
        public static Either<Exception,object> run(InputModel input) {

            Range c1 = (Range)input.dataSheet.Cells[input.fromRow, input.fromCol];
            Range c2 = (Range)input.dataSheet.Cells[input.toRow, input.toCol];
            Range inputRange =(Range) input.dataSheet.Range[c1, c2];

            MArray array = MArray.createFromRange(inputRange, input.groupByIndex, input.sortIndices);
            ImmutableList<MArrayWithGroup> either= array.splitByGroupFlag();
            foreach (MArrayWithGroup mg in either) {
                try {
                    mg.writeToSheet(input.fromCol);
                }catch(Exception e) {
                    return Either<Exception, object>.fail<Exception, object>(e);
                }
            }
             return Either<Exception,object>.NO_EXCEPTION;
        }
    }
}
