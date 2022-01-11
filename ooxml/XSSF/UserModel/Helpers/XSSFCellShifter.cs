/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using System.Collections.Generic;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using System;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Linq;
using NPOI.SS.UserModel.Helpers;

namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * @author Izumi Hoshino
     */
    public class XSSFCellShifter : CellShifter
    {

        public XSSFCellShifter(XSSFSheet sh)
            : base(sh)
        {
            sheet = sh;
        }

        public override void UpdateNamedRanges(FormulaShifter shifter)
        {
        }

        public override void UpdateFormulas(FormulaShifter shifter)
        {
        }

        private void UpdateSheetFormulas(ISheet sh, FormulaShifter Shifter)
        {
        }
        public override void UpdateRowFormulas(IRow row, FormulaShifter Shifter)
        {
        }

        private static String ShiftFormula(IRow row, String formula, FormulaShifter Shifter)
        {
            return String.Empty;
        }

        public override void UpdateConditionalFormatting(FormulaShifter Shifter)
        {
        }

        public override void UpdateHyperlinks(FormulaShifter shifter)
        {
        }

        private static CellRangeAddress ShiftRange(FormulaShifter Shifter, CellRangeAddress cra, int currentExternSheetIx)
        {
            throw new NotImplementedException();
        }

    }
}


