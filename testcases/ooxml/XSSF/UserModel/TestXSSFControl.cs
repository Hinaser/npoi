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
            //the sheet has one relationship and it is XSSFDrawing
            List<POIXMLDocumentPart.RelationPart> rels = sheet.RelationParts;
            Assert.AreEqual(1, rels.Count);
            POIXMLDocumentPart.RelationPart rp = rels[0];
            Assert.IsTrue(rp.DocumentPart is XSSFDrawing);

            XSSFDrawing drawing = (XSSFDrawing)rp.DocumentPart;
            //sheet.CreateDrawingPatriarch() should return the same instance of XSSFDrawing
            Assert.AreSame(drawing, sheet.CreateDrawingPatriarch());
            String drawingId = rp.Relationship.Id;

            //there should be a relation to this Drawing in the worksheet
            Assert.IsTrue(sheet.GetCTWorksheet().IsSetDrawing());
            Assert.AreEqual(drawingId, sheet.GetCTWorksheet().drawing.id);


            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(6, shapes.Count);

            Assert.IsTrue(shapes[(0)] is XSSFPicture);
            Assert.IsTrue(shapes[(1)] is XSSFPicture);
            Assert.IsTrue(shapes[(2)] is XSSFPicture);
            Assert.IsTrue(shapes[(3)] is XSSFPicture);
            Assert.IsTrue(shapes[(4)] is XSSFSimpleShape);
            Assert.IsTrue(shapes[(5)] is XSSFPicture);

            foreach (XSSFShape sh in shapes)
                Assert.IsNotNull(sh.GetAnchor());

            checkRewrite(wb);
            wb.Close();
        }
        private static void checkRewrite(XSSFWorkbook wb)
        {
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wb2);
            wb2.Close();
        }
    }
}