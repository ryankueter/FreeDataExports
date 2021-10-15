/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System.Collections.Generic;
using System.Xml.Linq;

namespace FreeDataExports.Spreadsheets.Ods1_3
{
    /// <summary>
    /// The worksheet class
    /// </summary>
    internal class Worksheet : IDataWorksheet
    {
        public Worksheet(string name)
        {
            Name = name;
            Rows = new List<IDataCell[]>();
            _columnWidths = new List<string>();
        }
        public string Name { get; } // Worksheet name
        public string TabColor { get; set; } // Worksheet tab color

        // A list of rows and cells
        internal List<IDataCell[]> Rows { get; set; }
        /// <summary>
        /// A method for adding rows
        /// </summary>
        /// <param name="c">Cells</param>
        public void AddRow(params IDataCell[] c)
        {
            Rows.Add(c);
        }

        /// <summary>
        /// Add a cell
        /// </summary>
        /// <param name="Data">The cell data.</param>
        /// <param name="Type">The cell's datatype.</param>
        /// <returns>Cell</returns>
        public IDataCell AddCell(object Data, DataType Type)
        {
            return new Cell(Data, Type);
        }

        // A list of column widths
        private List<string> _columnWidths { get; set; }
        /// <summary>
        /// A method for adding column widths
        /// </summary>
        /// <param name="columnWidths">Column widths</param>
        public void ColumnWidths(params string[] columnWidths)
        {
            for (int i = 0; i < columnWidths.Length; i++)
            {
                _columnWidths.Add(columnWidths[i]);
            }
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
