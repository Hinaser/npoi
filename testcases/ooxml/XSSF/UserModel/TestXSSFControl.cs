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
using NUnit.Framework;
using System;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Dml;
using NPOI.Util;
using System.Drawing;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Text;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI;
namespace TestCases.XSSF.UserModel
{
    /**
     * @author Izumi Hoshino
     */
    [TestFixture]
    public class TestXSSFControl
    {
        [Test]
        public void TestRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithControl.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            
            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(7, rels.Count);

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetExtControls());

            var controls = sheet.GetXSSFControls();
            for (int i = 0; i < rels.Count; i++)
            {
                var rp = rels[i];
                String rId = rp.Relationship.Id;
                if (rId != "rId4" && rId != "rId5" && rId != "rId6" && rId != "rId7"){
                    continue;
                }

                XSSFControl control = (XSSFControl)rp.DocumentPart;

                Assert.Contains(control, controls);
                Assert.IsTrue(rp.DocumentPart is XSSFControl);
                Assert.IsTrue(sheet.GetCTWorksheet().extControls.controls.Exists(c => c.id == rId));

                if (rId == "rId4")
                {
                    Assert.AreEqual(control.FormControlPr.objectType, "CheckBox");
                    control.FormControlPr.isChecked = "";
                }
                else if (rId == "rId5") Assert.AreEqual(control.FormControlPr.objectType, "Radio");
                else if (rId == "rId6") Assert.AreEqual(control.FormControlPr.objectType, "Radio");
                else if (rId == "rId7") Assert.AreEqual(control.FormControlPr.objectType, "Drop");
            }

            CheckRewrite(wb);
            wb.Close();
        }

        private static void CheckRewrite(XSSFWorkbook wb)
        {
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wb2);
            wb2.Close();
        }
    }
}