using System;
using System.Collections.Immutable;
using System.Linq;


namespace GroupingTool.model
{
    public class SortFlag : IComparable     {
        ImmutableList<IComparable> values;
        
        public SortFlag(ImmutableList<IComparable> values) {
            this.values = values;
        }


        public int CompareTo(object other) {
            if(other is SortFlag) {
                SortFlag o = (SortFlag)other;
                if(o.values.Count == this.values.Count) {
                    for (int x = 0; x < this.values.Count; ++x) {
                        IComparable ofThis = this.values[x];
                        IComparable ofOther = o.values[x];
                        int compareRs = ofThis.CompareTo(ofOther);
                        if (compareRs != 0) {
                            return compareRs;
                        }
                    }
                    return 0;
                } else {
                    throw new Exception("cannot compare");
                }
            } else {
                throw new Exception("cannot compare");
            }
        }
    }
}
