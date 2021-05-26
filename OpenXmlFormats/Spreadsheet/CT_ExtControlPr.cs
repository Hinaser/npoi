using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.ComponentModel;
using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class Col
    {
        public int col;

        public static Col Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            Col ctObj = new Col();
            ctObj.col = int.Parse(node.InnerText);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}>", nodeName));
            sw.Write(string.Format("{0}", this.col));
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class ColOff
    {
        public int colOff;

        public static ColOff Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            ColOff ctObj = new ColOff();
            ctObj.colOff = int.Parse(node.InnerText);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}>", nodeName));
            sw.Write(string.Format("{0}", this.colOff));
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class Row
    {
        public int row;

        public static Row Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            Row ctObj = new Row();
            ctObj.row = int.Parse(node.InnerText);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}>", nodeName));
            sw.Write(string.Format("{0}", this.row));
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class RowOff
    {
        public int rowOff;

        public static RowOff Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            RowOff ctObj = new RowOff();
            ctObj.rowOff = int.Parse(node.InnerText);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}>", nodeName));
            sw.Write(string.Format("{0}", this.rowOff));
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ExtCellPosition
    {
        private Col colField;
        private ColOff colOffField;
        private Row rowField;
        private RowOff rowOffField;

        public Col col
        {
            get
            {
                return colField;
            }
            set
            {
                this.colField = value;
            }
        }

        public ColOff colOff
        {
            get
            {
                return colOffField;
            }
            set
            {
                this.colOffField = value;
            }
        }

        public Row row
        {
            get
            {
                return rowField;
            }
            set
            {
                this.rowField = value;
            }
        }

        public RowOff rowOff
        {
            get
            {
                return rowOffField;
            }
            set
            {
                this.rowOffField = value;
            }
        }
        public static CT_ExtCellPosition Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExtCellPosition ctObj = new CT_ExtCellPosition();
            foreach(XmlNode n in node.ChildNodes)
            {
                if(n.LocalName == "col")
                {
                    ctObj.col = Col.Parse(n, namespaceManager);
                }
                else if(n.LocalName == "colOff")
                {
                    ctObj.colOff = ColOff.Parse(n, namespaceManager);
                }
                else if (n.LocalName == "row")
                {
                    ctObj.row = Row.Parse(n, namespaceManager);
                }
                else if (n.LocalName == "rowOff")
                {
                    ctObj.rowOff = RowOff.Parse(n, namespaceManager);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}>", nodeName));
            this.col.Write(sw, "col");
            this.colOff.Write(sw, "colOff");
            this.row.Write(sw, "row");
            this.rowOff.Write(sw, "rowOff");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExtAnchor
    {
        private int moveWithCellsField;
        private CT_ExtCellPosition fromField;
        private CT_ExtCellPosition toField;

        public int moveWithCells
        {
            get
            {
                return moveWithCellsField;
            }
            set
            {
                this.moveWithCellsField = value;
            }
        }

        public CT_ExtCellPosition from
        {
            get
            {
                return fromField;
            }
            set
            {
                this.fromField = value;
            }
        }

        public CT_ExtCellPosition to
        {
            get
            {
                return toField;
            }
            set
            {
                this.toField = value;
            }
        }

        public static CT_ExtAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExtAnchor ctObj = new CT_ExtAnchor();
            ctObj.moveWithCells = XmlHelper.ReadInt(node.Attributes["moveWithCells"]);
            foreach (XmlNode n in node.ChildNodes)
            {
                if(n.LocalName == "from")
                {
                    ctObj.from = CT_ExtCellPosition.Parse(n, namespaceManager);
                }
                else if (n.LocalName == "to")
                {
                    ctObj.to = CT_ExtCellPosition.Parse(n, namespaceManager);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "moveWithCells", this.moveWithCells);
            sw.Write(">");
            this.from.Write(sw, "from");
            this.to.Write(sw, "to");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_ExtControlPr
    {

        private uint defaultSizeField;

        private int autoFillField;

        private int autoLineField;

        private int autoPictField;

        private CT_ExtAnchor anchorField;

        public uint defaultSize
        {
            get
            {
                return this.defaultSizeField;
            }
            set
            {
                this.defaultSizeField = value;
            }
        }

        public int autoFill
        {
            get
            {
                return this.autoFillField;
            }
            set
            {
                this.autoFillField = value;
            }
        }

        public int autoLine
        {
            get
            {
                return this.autoLineField;
            }
            set
            {
                this.autoLineField = value;
            }
        }

        public int autoPict
        {
            get
            {
                return this.autoPictField;
            }
            set
            {
                this.autoPictField = value;
            }
        }

        public CT_ExtAnchor anchor
        {
            get
            {
                return anchorField;
            }
            set
            {
                this.anchorField = value;
            }
        }

        public static CT_ExtControlPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExtControlPr ctObj = new CT_ExtControlPr();
            ctObj.defaultSize = XmlHelper.ReadUInt(node.Attributes["defaultSize"]);
            ctObj.autoFill = XmlHelper.ReadInt(node.Attributes["autoFill"]);
            ctObj.autoLine = XmlHelper.ReadInt(node.Attributes["autoLine"]);
            ctObj.autoPict = XmlHelper.ReadInt(node.Attributes["autoPict"]);
            foreach(XmlNode n in node.ChildNodes)
            {
                if(n.LocalName == "anchor")
                {
                    ctObj.anchor = CT_ExtAnchor.Parse(n, namespaceManager);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "defaultSize", this.defaultSize);
            XmlHelper.WriteAttribute(sw, "autoFill", this.autoFill);
            XmlHelper.WriteAttribute(sw, "autoLine", this.autoLine);
            XmlHelper.WriteAttribute(sw, "autoPict", this.autoPict);
            sw.Write(">");
            this.anchor.Write(sw, "anchor");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
}
