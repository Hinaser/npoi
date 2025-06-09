/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.SS.Formula.Eval;
using System;
using System.Text;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using System.Globalization;
using NPOI.Util;
using System.Collections.Generic;
namespace NPOI.SS.Formula.Functions
{

    /**
     * Implementation for Excel Unicode() function.<br/>
     * <br/>
     * <b>Syntax</b>:<br/> <b>Unicode  </b>(<b>string</b>)<br/>
     * <br/>
     * Returns the number (code point) corresponding to the first character of the text.
     *
     * @author iz hyphen hoshino @ toyokeizai dot co dot jp
     */
    public class Unicode : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new Unicode();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval strVE)
        {
            String s;
            if (strVE is RefEval)
            {
                RefEval re = (RefEval)strVE;
                s = OperandResolver.CoerceValueToString(re.GetInnerValueEval(re.FirstSheetIndex));
            }
            else
            {
                s = OperandResolver.CoerceValueToString(strVE);
            }

            if (s.Length < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                Int32 codePoint = Char.ConvertToUtf32(s, 0);
                return new NumberEval(codePoint);
            }
            catch (ArgumentException)
            {
                // 不正なサロゲートペアの場合
                return ErrorEval.VALUE_INVALID;
            }
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]);
        }
    }
}