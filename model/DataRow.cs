using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
namespace GroupingTool.model {
	/**
	 * Represent a row
	 */
	public class DataRow : IComparable<DataRow> {

		public readonly object groupFlag;
		public readonly SortFlag sortFlag;
		public readonly object[] dataArray;

		public DataRow(object[] data, int groupByIndex, List<int> sortFlagIndices) {
			this.dataArray = data;
			this.groupFlag = data[groupByIndex];
			List<IComparable> sfs = new List<IComparable>();
			foreach (int i in sortFlagIndices) {
				sfs.Add((IComparable)data[i]);
            }
			this.sortFlag = new SortFlag(ImmutableList.Create<IComparable>(sfs.ToArray()));

		}

        public int CompareTo(DataRow obj) {
			return this.sortFlag.CompareTo(obj.sortFlag);
        }

		public int size() {
			return this.dataArray.Count();
        }
    }
}
