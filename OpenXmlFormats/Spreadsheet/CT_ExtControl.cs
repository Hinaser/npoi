namespace NPOI.OpenXmlFormats.Spreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml.Serialization;
    using System.Xml;
    using NPOI.OpenXml4Net.Util;
    using NPOI.OpenXml4Net.OPC;

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExtControl
    {
        private uint shapeIdField;

        private string idField;

        private string nameField;

        private CT_ExtControlPr controlPrField;

        public uint shapeId
        {
            get
            {
                return this.shapeIdField;
            }
            set
            {
                this.shapeIdField = value;
            }
        }

        // TODO is the following correct?
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        public CT_ExtControlPr controlPr
        {
            get
            {
                return this.controlPrField;
            }
            set
            {
                this.controlPrField = value;
            }
        }

        public static CT_ExtControl Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExtControl ctObj = new CT_ExtControl();
            ctObj.shapeId = XmlHelper.ReadUInt(node.Attributes["shapeId"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id", PackageNamespaces.SCHEMA_RELATIONSHIPS]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "controlPr")
                    ctObj.controlPr = CT_ExtControlPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "shapeId", this.shapeId);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            XmlHelper.WriteAttribute(sw, "name", this.name);
            sw.Write(">");
            this.controlPr.Write(sw, "controlPr");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExtControls
    {

        private List<CT_ExtControl> controlsField;

        public List<CT_ExtControl> controls
        {
            get
            {
                return this.controlsField;
            }
            set
            {
                this.controlsField = value;
            }
        }

        public static CT_ExtControls Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExtControls ctObj = new CT_ExtControls();
            ctObj.controls = new List<CT_ExtControl>();

            foreach (XmlNode n in node)
            {
                if (n.LocalName == "AlternateContent"
                    && n.NamespaceURI == "http://schemas.openxmlformats.org/markup-compatibility/2006"
                    && n.ChildNodes.Count > 0
                    && n.ChildNodes[0].LocalName == "Choice"
                    && n.ChildNodes[0].NamespaceURI == "http://schemas.openxmlformats.org/markup-compatibility/2006"
                    && n.ChildNodes[0].ChildNodes.Count > 0
                )
                {
                    var cn2 = n.ChildNodes[0].ChildNodes[0];
                    if (cn2.LocalName == "control")
                    {
                        ctObj.controls.Add(CT_ExtControl.Parse(cn2, namespaceManager));
                    }
                }
                else if(n.LocalName == "control")
                {
                    ctObj.controls.Add(CT_ExtControl.Parse(n, namespaceManager));
                }
            }

            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
            sw.Write("<mc:Choice Requires=\"x14\">");
            sw.Write(string.Format("<{0}>", nodeName));
            if (this.controls != null)
            {
                foreach (CT_ExtControl control in this.controls)
                {
                    sw.Write("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
                    sw.Write("<mc:Choice Requires=\"x14\">");
                    control.Write(sw, "control");
                    sw.Write("</mc:Choice>");
                    sw.Write("</mc:AlternateContent>");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
            sw.Write("</mc:Choice>");
            sw.Write("</mc:AlternateContent>");
        }
    }
}
