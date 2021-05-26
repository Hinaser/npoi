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

using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.Spreadsheet; // http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System.Collections.Generic;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Util;
using NPOI.Util;


namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a SpreadsheetML Control
     *
     * @author Izumi Hoshino
     */
    public class XSSFControl : POIXMLDocumentPart
    {
        public static String NAMESPACE_A = XSSFRelation.NS_DRAWINGML;
        public static String NAMESPACE_C = XSSFRelation.NS_CHART;

        /**
         * Create a new SpreadsheetML Control
         *
         */
        public XSSFControl()
            : base()
        {
        }

        /**
         * Constructor called when read data from a file.
         */
        internal XSSFControl(PackagePart part)
            : base(part)
        {
            XmlDocument xmldoc = ConvertStreamToXml(part.GetInputStream());
            // drawing = CT_Drawing.Parse(xmldoc, NamespaceManager);
            int a = 0;
        }
        public XSSFControl(PackagePart part, PackageRelationship rel)
            : this(part)
        {
        }
            
        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            // drawing.Save(out1);
            out1.Close();
        }

    }
}

