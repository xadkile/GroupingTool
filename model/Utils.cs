using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class Utils {
        private Utils() { }
        // stole from stackoverflow. too lazy to write one
        public static T[,] To2D<T>(T[][] source) {
            try {
                int FirstDim = source.Length;
                int SecondDim = source.GroupBy(row => row.Length).Single().Key; // throws InvalidOperationException if source is not rectangular

                var result = new T[FirstDim, SecondDim];
                for (int i = 0; i < FirstDim; ++i)
                    for (int j = 0; j < SecondDim; ++j)
                        result[i, j] = source[i][j];

                return result;
            } catch (InvalidOperationException) {
                throw new InvalidOperationException("The given jagged array is not rectangular.");
            }
        }

        public static object[] processNullArray(object[] inputArr) {
            List<object> tmpList = new List<object>();
            foreach(object e in inputArr) {
                if(e == null) {
                    tmpList.Add("");
                } else {
                    tmpList.Add(e);
                }
            }
            return tmpList.ToArray();
        }

        public static T[][] toJaggedArray<T>(T[,] twoDimensionalArray, bool isStartAtOne) {
            int offset = 0;
            if (isStartAtOne) {
                offset = 1;
            }

            int rowsFirstIndex = twoDimensionalArray.GetLowerBound(0);
            int rowsLastIndex = twoDimensionalArray.GetUpperBound(0);
            int numberOfRows = rowsLastIndex - rowsFirstIndex + 1;

            int columnsFirstIndex = twoDimensionalArray.GetLowerBound(1);
            int columnsLastIndex = twoDimensionalArray.GetUpperBound(1);
            int numberOfColumns = columnsLastIndex - columnsFirstIndex + 1;

            T[][] jaggedArray = new T[numberOfRows][];
            for (int r = rowsFirstIndex; r <= rowsLastIndex; ++r) {
                jaggedArray[r-offset] = new T[numberOfColumns];

                for (int c = columnsFirstIndex; c <= columnsLastIndex; ++c) {
                    jaggedArray[r-offset][c-offset] = twoDimensionalArray[r, c];
                }
            }
            return jaggedArray;
        }

        static public Range getLastRange(Worksheet sheet) {
            Range last = (Range)sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            return last;
        }
        static public int getLastRow(Worksheet sheet) {
            Range last = getLastRange(sheet);
            //Range used = (Range)sheet.UsedRange;
            if (last.Row == 1 && last.Column==1 && last.Value2 == null) {
                return 0;
            } else {
                //Range last = getLastRange(sheet);
                // get last row count

                int nextRow = last.Row;
                return nextRow;
            }
        }

        static public int getFirstRow(Worksheet sheet) {
            return ((Range)sheet.UsedRange).Row;
        }
        
    }
}
