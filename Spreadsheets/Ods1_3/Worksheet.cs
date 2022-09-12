/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace FreeDataExports.Spreadsheets.Ods1_3
{
    /// <summary>
    /// The worksheet class
    /// </summary>
    internal sealed class Worksheet : IDataWorksheet
    {
        public Worksheet(string name)
        {
            Name = name;
            Rows = new List<Row>();
            CurrentRow = new Row();
            _columnWidths = new List<string>();
        }
        public string Name { get; } // Worksheet name
        public string TabColor { get; set; } // Worksheet tab color

        // Stores a list of rows
        internal List<Row> Rows { get; set; }

        // Stores the current row being loaded
        internal Row CurrentRow { get; set; }
        public IDataWorksheet AddRow()
        {
            CurrentRow = new Row();
            Rows.Add(CurrentRow);
            return this;
        }

        public IDataWorksheet AddCell(object Data, DataType Type)
        {
            CurrentRow.Add(new Cell(Data, Type));
            return this;
        }

        // A list of column widths
        private List<string> _columnWidths { get; set; }

        /// <summary>
        /// A method for adding column widths
        /// </summary>
        /// <param name="columnWidths">Column widths</param>
        public void ColumnWidths(params string[] columnWidths)
        {
            _columnWidths.AddRange(columnWidths);
        }

        internal List<string> GetColumnWidths()
        {
            return _columnWidths;
        }

        internal int GetColumnCount()
        {
            return _columnWidths.Count;
        }

        /// <summary>
        /// Worksheet data
        /// </summary>
        /// <returns></returns>
        public XElement Get_config_item_map_entry()
        {
            XNamespace config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

            var config_item_map_entry = new XElement(config + "config-item-map-entry", new XAttribute(config + "name", Name),
                new XElement(config + "config-item", new XAttribute(config + "name", "CursorPositionX"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "CursorPositionY"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "HorizontalSplitMode"), new XAttribute(config + "type", "short"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "VerticalSplitMode"), new XAttribute(config + "type", "short"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "HorizontalSplitPosition"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "VerticalSplitPosition"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "ActiveSplitRange"), new XAttribute(config + "type", "short"), 2),
                new XElement(config + "config-item", new XAttribute(config + "name", "PositionLeft"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "PositionRight"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "PositionTop"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "PositionBottom"), new XAttribute(config + "type", "int"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "ZoomType"), new XAttribute(config + "type", "short"), 0),
                new XElement(config + "config-item", new XAttribute(config + "name", "ZoomValue"), new XAttribute(config + "type", "int"), 100),
                new XElement(config + "config-item", new XAttribute(config + "name", "PageViewZoomValue"), new XAttribute(config + "type", "int"), 60),
                new XElement(config + "config-item", new XAttribute(config + "name", "ShowGrid"), new XAttribute(config + "type", "boolean"), true),
                new XElement(config + "config-item", new XAttribute(config + "name", "AnchoredTextOverflowLegacy"), new XAttribute(config + "type", "boolean"), false));

            return config_item_map_entry;
        }
    }
}
