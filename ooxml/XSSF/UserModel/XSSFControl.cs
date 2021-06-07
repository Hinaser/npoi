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
using NPOI.OpenXml4Net.Util;


namespace NPOI.XSSF.UserModel
{
    public class CT_FormControlPr
    {
        private string objectTypeField; // Checkbox, Radio, Drop
        private string isCheckedField; // Original attr name is 'checked'. Checkbox
        private string fmlaLinkField; // Checkbox, Radio, Drop
        private string lockTextField; // Checkbox, Radio
        private string noThreeDField; // Checkbox, Radio, Drop
        private string firstButtonField; // Radio
        private string dropStyleField; // Drop
        private string dxField; // Drop
        private string fmlaRangeField; // Drop
        private string selField; // Drop
        private string valField; // Drop

        public string objectType
        {
            get { return objectTypeField; }
            set { objectTypeField = value;  }
        }

        public string isChecked
        {
            get { return isCheckedField; }
            set { isCheckedField = value; }
        }

        public string fmlaLink
        {
            get { return fmlaLinkField; }
            set { fmlaLinkField = value; }
        }

        public string lockText
        {
            get { return lockTextField; }
            set { lockTextField = value; }
        }

        public string noThreeD
        {
            get { return noThreeDField; }
            set { noThreeDField = value; }
        }

        public string firstButton
        {
            get { return firstButtonField; }
            set { firstButtonField = value; }
        }

        public string dropStyle
        {
            get { return dropStyleField; }
            set { dropStyleField = value; }
        }

        public string dx
        {
            get { return dxField; }
            set { dxField = value; }
        }

        public string fmlaRange
        {
            get { return fmlaRangeField; }
            set { fmlaRangeField = value; }
        }

        public string sel
        {
            get { return selField; }
            set { selField = value; }
        }

        public string val
        {
            get { return valField; }
            set { valField = value; }
        }

        public static CT_FormControlPr Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceManager)
        {
            if (xmldoc == null)
                return null;
            CT_FormControlPr ctObj = new CT_FormControlPr();

            XmlNode node = xmldoc.SelectSingleNode("/*[local-name()='formControlPr']", namespaceManager);

            ctObj.objectType = XmlHelper.ReadString(node.Attributes["objectType"]);
            if (ctObj.objectType == "CheckBox")
            {
                ctObj.isChecked = XmlHelper.ReadString(node.Attributes["checked"]);
                ctObj.fmlaLink = XmlHelper.ReadString(node.Attributes["fmlaLink"]);
                ctObj.lockText = XmlHelper.ReadString(node.Attributes["lockText"]);
                ctObj.noThreeD = XmlHelper.ReadString(node.Attributes["noThreeD"]);
            }
            else if(ctObj.objectType == "Radio")
            {
                ctObj.firstButton = XmlHelper.ReadString(node.Attributes["firstButton"]);
                ctObj.fmlaLink = XmlHelper.ReadString(node.Attributes["fmlaLink"]);
                ctObj.lockText = XmlHelper.ReadString(node.Attributes["lockText"]);
                ctObj.noThreeD = XmlHelper.ReadString(node.Attributes["noThreeD"]);
            }
            else if(ctObj.objectType == "Drop")
            {
                ctObj.dropStyle = XmlHelper.ReadString(node.Attributes["dropStyle"]);
                ctObj.dx = XmlHelper.ReadString(node.Attributes["dx"]);
                ctObj.fmlaLink = XmlHelper.ReadString(node.Attributes["fmlaLink"]);
                ctObj.fmlaRange = XmlHelper.ReadString(node.Attributes["fmlaRange"]);
                ctObj.noThreeD = XmlHelper.ReadString(node.Attributes["noThreeD"]);
                ctObj.sel = XmlHelper.ReadString(node.Attributes["val"]);
                ctObj.val = XmlHelper.ReadString(node.Attributes["val"]);
            }

            return ctObj;
        }

        internal void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");

                sw.Write(string.Format("<{0}", "formControlPr"));
                sw.Write(" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"");
                XmlHelper.WriteAttribute(sw, "objectType", objectType);

                if (!String.IsNullOrEmpty(isChecked)) XmlHelper.WriteAttribute(sw, "checked", isChecked);
                if (!String.IsNullOrEmpty(fmlaLink)) XmlHelper.WriteAttribute(sw, "fmlaLink", fmlaLink);
                if (!String.IsNullOrEmpty(lockText)) XmlHelper.WriteAttribute(sw, "lockText", lockText);
                if (!String.IsNullOrEmpty(noThreeD)) XmlHelper.WriteAttribute(sw, "noThreeD", noThreeD);
                if (!String.IsNullOrEmpty(firstButton)) XmlHelper.WriteAttribute(sw, "firstButton", firstButton);
                if (!String.IsNullOrEmpty(dropStyle)) XmlHelper.WriteAttribute(sw, "dropStyle", dropStyle);
                if (!String.IsNullOrEmpty(dx)) XmlHelper.WriteAttribute(sw, "dx", dx);
                if (!String.IsNullOrEmpty(fmlaRange)) XmlHelper.WriteAttribute(sw, "fmlaRange", fmlaRange);
                if (!String.IsNullOrEmpty(sel)) XmlHelper.WriteAttribute(sw, "sel", sel);
                if (!String.IsNullOrEmpty(val)) XmlHelper.WriteAttribute(sw, "val", val);

                sw.Write(" />");
            }
        }
    }

    /**
     * Represents a SpreadsheetML Control
     *
     * @author Izumi Hoshino
     */
    public class XSSFControl : POIXMLDocumentPart
    {
        private CT_FormControlPr formControlPrField = new CT_FormControlPr();

        public CT_FormControlPr FormControlPr
        {
            get { return formControlPrField; }
        }

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
            formControlPrField = CT_FormControlPr.Parse(xmldoc, NamespaceManager);
        }
        public XSSFControl(PackagePart part, PackageRelationship rel)
            : this(part)
        {
        }
            
        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            formControlPrField.Save(out1);
            out1.Close();
        }
    }
}

