using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Immutable;
using Microsoft.Office.Interop.Excel;

namespace GroupingTool.model {
    public class MArray {
        public readonly ImmutableList<DataRow> rows;

        public MArray(ImmutableList<DataRow> rows) {
            this.rows = rows;
        }

        public static MArray createFromRange(Range range, int groupByIndex, List<int> sortFlagIndices) {

            List<DataRow> dataRowList = new List<DataRow>();
            foreach (Range rowRange in range.Rows) {
                object[,] valueArray= (object[,])rowRange.Value2;
                object[][] jaggedValueArray = Utils.toJaggedArray(valueArray, true);
                object[] rowData = jaggedValueArray[0];
                dataRowList.Add(new DataRow(rowData, groupByIndex, sortFlagIndices));
            }
            return new MArray(dataRowList.ToImmutableList());
        }

        public MArray sortRows() {
            return new MArray(this.rows.Sort());
        }

        /**
         * pour data into a sheet with the top left corner being defined by [row] and [col]
         */
        public void pourToSheet(Worksheet sheet, int row,int col) {
            Range c1 =(Range) sheet.Cells[row, col];
            Range c2 = (Range)sheet.Cells[row + this.rows.Count-1, col+this.rows[0].size()-1];
            Range target = (Range)sheet.Range[c1, c2];
            target.Value2 = this.makeDataArray();
        }
        
        /**
         * transform data in this MArray to a 2d array
         */
        private object[,] makeDataArray() {
            var rt = this.rows.Select(e => e.dataArray).ToArray();
            return Utils.To2D(rt);
        }

        public ImmutableList<MArrayWithGroup> splitByGroupFlag() {
            List<MArrayWithGroup> rt = new List<MArrayWithGroup>();
            IEnumerable<IGrouping<object, DataRow>> groups = this.rows.GroupBy((DataRow e )=> e.groupFlag);
            foreach (IGrouping<object,DataRow> grp in groups) {
                rt.Add(new MArrayWithGroup(
                    grp.Key, 
                    new MArray(grp.ToImmutableList()).sortRows()
                    ));
            }
            return rt.ToImmutableList();
        }
    }
}
