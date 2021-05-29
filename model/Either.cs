using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupingTool.model {
    public class  Either<L, R> {
        public readonly L left;
        public readonly R right;
        public Either(L left, R right) {
            this.left = left;
            this.right = right;
        }

        public static Either<Exception, object> NO_EXCEPTION = new Either<Exception, object>(null, null);

        public static Either<LX,RX> from<LX,RX>(LX left, RX right) where LX : Exception {
            return new Either<LX,RX>(left, right);
        }

        public static Either<LX, RX> ok<LX, RX>(RX right) where LX :Exception  {
            return new Either<LX, RX>(null, right);
        }

        public static Either<LX,RX> fail<LX, RX>(LX left) where LX : Exception {
            return new Either<LX, RX>(left, default(RX));
        }

        public bool hasLeft() {
            return this.left != null;
        }

        public bool hasRight() {
            return this.right != null;
        }

        public bool isOk() {
            return !this.hasException() && this.hasRight();
        }

        public bool hasException() {
            return this.hasLeft() && this.left is Exception;
        }

        public R get() {
            return this.right;
        }

        public L getException() {
            return this.left;
        }
        
        public bool notHasException() {
            return this.left == null;
        }
    }
}
