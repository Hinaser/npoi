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

using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;

namespace NPOI.OOXML.Testcases.XSSF.Formula.Functions
{
    /**
     * @author Izumi Hoshino
     */
    [TestFixture]
    public class TestFunctionUnicode
    {
        [Test]
        public void TestEval()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("FunctionUnicode.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            var cellA1 = sheet.GetRow(0).Cells[0];
            var cellB1 = sheet.GetRow(0).Cells[1];

            Assert.AreEqual(cellA1.StringCellValue, "あ");
            Assert.AreEqual(cellB1.NumericCellValue, 12354);

            cellA1.SetCellValue("い");

            Assert.DoesNotThrow(() => { XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb); });

            Assert.AreEqual(cellA1.StringCellValue, "い");
            Assert.AreEqual(cellB1.NumericCellValue, 12356);

            cellA1.SetCellValue("a");

            Assert.DoesNotThrow(() => { XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb); });

            Assert.AreEqual(cellA1.StringCellValue, "a");
            Assert.AreEqual(cellB1.NumericCellValue, 97);

            wb.Close();
        }
    }
}
