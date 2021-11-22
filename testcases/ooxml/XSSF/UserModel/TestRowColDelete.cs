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

using TestCases.SS.UserModel;
using NUnit.Framework;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.Text;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Collections;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;
using System.IO;


namespace TestCases.XSSF.UserModel
{

    [TestFixture]
    public class TestRowColDelete : BaseTestXSheet
    {

        public TestRowColDelete()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void DeletePartialRowCol()
        {
            XSSFWorkbook wb = XSSFITestDataProvider.instance.OpenSampleWorkbook("RowColDeleting.xlsx") as XSSFWorkbook;
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            // WriteToFile(wb);
            wb.Close();
        }

        private static void WriteToFile(XSSFWorkbook wb)
        {
            using (var fs = File.Create(@"C:\Users\Hinaser\Desktop\bbb.xlsx"))
            {
                wb.Write(fs);
            }
        }
    }
}

