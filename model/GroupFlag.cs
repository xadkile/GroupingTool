using System;

namespace GroupingTool.model {
	public class GroupFlag {
		object value;
		public GroupFlag(IComparable value) {
			this.value = value;
		}

		public override bool Equals(object obj) {
			if (obj is GroupFlag) {
				return this.value.Equals(((GroupFlag)obj).value);
			} else {
				return false;
			}
		}
	}
}
