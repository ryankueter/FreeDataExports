﻿/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace FreeDataExports.Spreadsheets.XL2019
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
            SharedStrings = new OrderedDictionary();
            _columnWidths = new List<string>();
        }
        internal int Id { get; set; } // Worksheet Id
        public string Name { get; } // Worksheet name
        public string TabColor { get; set; } // Worksheet tab color

        // Stores a copy of the SharedStrings from the workbook for index information
        internal OrderedDictionary SharedStrings { get; set; }
        
        // Stores a list of data definitions for the cell data
        internal List<DataDefinition> CellFormats = new List<DataDefinition>();

        // Stores a list of rows
        internal List<IDataCell[]> Rows { get; set; }
        /// <summary>
        /// A method for adding cells to a row
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

        // Stores a list of column widths
        private List<string> _columnWidths { get; set; }
        /// <summary>
        /// A method for supplying column widths
        /// </summary>
        /// <param name="columnWidths">Column Widths</param>
        public void ColumnWidths(params string[] columnWidths)
        {
            foreach (var d in columnWidths)
            {
                _columnWidths.Add(d);
            }
        }

        // Some XML Namespaces used by the worksheet
        private readonly XNamespace xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private readonly XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private readonly XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private readonly XNamespace x14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
        private readonly XNamespace xr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
        private readonly XNamespace xr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
        private readonly XNamespace xr3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";

        /// <summary>
        /// The worksheet data
        /// </summary>
        /// <param name="sharedStrings">Stores deduplicated strings.</param>
        /// <returns></returns>
        internal XDocument xl_worksheets_sheet(OrderedDictionary sharedStrings)
        {
            SharedStrings = sharedStrings;
            int width = GetColumnCount();

            var worksheet = new XElement("worksheet",
                    new XAttribute("xmlns", xmlns),
                    new XAttribute(XNamespace.Xmlns + "r", r),
                    new XAttribute(XNamespace.Xmlns + "mc", mc),
                    new XAttribute(mc + "Ignorable", "x14ac xr xr2 xr3"),
                    new XAttribute(XNamespace.Xmlns + "x14ac", x14ac),
                    new XAttribute(XNamespace.Xmlns + "xr", xr),
                    new XAttribute(XNamespace.Xmlns + "xr2", xr2),
                    new XAttribute(XNamespace.Xmlns + "xr3", xr3),
                    new XAttribute(xr + "uid", Utilities.NewGuid()));

            // Add the tab color if specified
            if (String.IsNullOrEmpty(TabColor) == false)
            {
                worksheet.Add(new XElement("sheetPr", new XElement("tabColor", new XAttribute("rgb", TabColor))));
            }
            worksheet.Add(GetDimension());
            worksheet.Add(new XElement("sheetViews", new XElement("sheetView", new XAttribute("workbookViewId", "0"))));
            worksheet.Add(new XElement("sheetFormatPr", new XAttribute("defaultRowHeight", "15"), new XAttribute(x14ac + "dyDescent", "0.25")));

            if (_columnWidths.Count > 0)
            {
                var cols = new XElement("cols");

                int i = 1; // Column range
                foreach (var c in _columnWidths)
                {
                    cols.Add(new XElement("col",
                                        new XAttribute("min", i.ToString()),
                                        new XAttribute("max", i.ToString()),
                                        new XAttribute("width", c),
                                        new XAttribute("customWidth", "1")));
                    i++;
                }
                worksheet.Add(cols);
            }
            worksheet.Add(GetSheetData());
            worksheet.Add(new XElement("pageMargins",
                        new XAttribute("left", "0.7"),
                        new XAttribute("right", "0.7"),
                        new XAttribute("top", "0.75"),
                        new XAttribute("bottom", "0.75"),
                        new XAttribute("header", "0.3"),
                        new XAttribute("footer", "0.3")));
                        
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), worksheet);
        }

        /// <summary>
        /// Worksheet relationships
        /// </summary>
        /// <returns></returns>
        internal XDocument xl_worksheets_rels_sheet()
        {
            XNamespace xmlns = "http://schemas.openxmlformats.org/package/2006/relationships";

            var worksheet = new XElement("Relationships", new XAttribute("xmlns", xmlns));
                        
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), worksheet);
        }

        // Gets the cell ranges
        private XElement GetDimension()
        {
            if (Rows.Count > 0)
            {
                int width = GetColumnCount();
                int depth = GetRowCount();
                string columnRow = Utilities.GetIndex(width) + depth.ToString();

                if (String.IsNullOrEmpty(columnRow) == false)
                {
                    return new XElement("dimension", new XAttribute("ref", $"A1:{columnRow}"));
                }
            }

            return new XElement("dimension", new XAttribute("ref", "A1"));
        }

        /// <summary>
        /// Get the sheet data if any exists
        /// </summary>
        /// <returns></returns>
        private XElement GetSheetData()
        {
            int depth = GetRowCount();
            if (depth == 0)
            {
                return new XElement("sheetData");
            }
            else
            {
                var sheetData = new XElement("sheetData");
                return GetRows(sheetData);
            }
        }

        /// <summary>
        /// Get the rows for the sheet data
        /// </summary>
        /// <param name="sheetData"></param>
        /// <returns></returns>
        private XElement GetRows(XElement sheetData)
        {
            if (Rows.Count > 0)
            {
                int width = GetColumnCount();

                // Add the rows
                int r = 1; // Current row
                foreach (var o in Rows)
                {
                    var row = new XElement("row", new XAttribute("r", r), new XAttribute("spans", $"{1}:{width}"), new XAttribute(x14ac + "dyDescent", "0.25"));

                    int col = 1;
                    // Foreach cell in the row
                    foreach (var c in o)
                    {
                        // Strings, numbers, and booleans are different from other datatypes
                        if (c.DataType == DataType.String)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "s"), new XElement("v", SharedStrings[c.Value])));
                        }
                        else if (c.DataType == DataType.Number)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XElement("v", c.Value)));
                        }
                        if (c.DataType == DataType.Boolean)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "b"), new XElement("v", c.Value)));
                        }
                        else
                        {
                            foreach (var n in CellFormats)
                            {
                                if (n.type == (int)c.DataType)
                                {
                                    row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("s", n.index.ToString()), new XElement("v", c.Value)));
                                }
                            }
                        }
                        col++;
                    }
                    r++;

                    sheetData.Add(row);
                }
            }

            return sheetData;
        }

        private int GetColumnHeaderCount()
        {
            int width = 0;
            foreach (var w in Rows)
            {
                width = w.Length;
                break;
            }
            return width;
        }

        private int GetColumnCount()
        {
            int count = 0;
            foreach (var r in Rows)
            {
                count = r.Length;
                break;
            }
            return count;
        }

        private int GetRowCount()
        { 
            return Rows.Count;
        }
    }
}