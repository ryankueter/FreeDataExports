﻿/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Runtime.InteropServices;

namespace FreeDataExports.Spreadsheets.XL2019
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
            SharedStrings = new OrderedDictionary();
            CurrentRow = new Row();
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

        // Stores a list of column widths
        private List<string> _columnWidths { get; set; }

        /// <summary>
        /// A method for supplying column widths
        /// </summary>
        /// <param name="columnWidths">Column Widths</param>
        public void ColumnWidths(params string[] columnWidths)
        {
            _columnWidths.AddRange(columnWidths);
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

                // Column range
                for (int i = 0; i < _columnWidths.Count; i++)
                {
                    int w = i + 1;
                    cols.Add(new XElement("col",
                                         new XAttribute("min", w.ToString()),
                                         new XAttribute("max", w.ToString()),
                                         new XAttribute("width", _columnWidths[i]),
                                         new XAttribute("customWidth", "1")));
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
                int depth = Rows.Count;
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
            int depth = Rows.Count;
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

#if NET6_0_OR_GREATER
                // Add the rows
                var rows = CollectionsMarshal.AsSpan(Rows);
                for (int i = 0; i < rows.Length; i++)
                {
                    int r = i + 1; // Current row
                    var row = new XElement("row", new XAttribute("r", r), new XAttribute("spans", $"{1}:{width}"), new XAttribute(x14ac + "dyDescent", "0.25"));

                    // Iterate the cells
                    for (int c = 0; c < rows[i].Count; c++)
                    {
                        int col = c + 1;

                        // Strings, numbers, and booleans are different from other datatypes
                        if (rows[i][c].DataType == DataType.String)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "s"), new XElement("v", SharedStrings[rows[i][c].Value])));
                        }
                        else if (rows[i][c].DataType == DataType.Number)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XElement("v", rows[i][c].Value)));
                        }
                        if (rows[i][c].DataType == DataType.Boolean)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "b"), new XElement("v", rows[i][c].Value)));
                        }
                        else
                        {
                            for (int n = 0; n < CellFormats.Count; n++)
                            {
                                if (CellFormats[n].type == (int)rows[i][c].DataType)
                                {
                                    row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("s", CellFormats[n].index.ToString()), new XElement("v", rows[i][c].Value)));
                                }
                            }
                        }
                    }

                    sheetData.Add(row);
                }
#elif NETSTANDARD2_0_OR_GREATER
                // Add the rows
                for (int i = 0; i < Rows.Count; i++)
                {
                    int r = i + 1; // Current row
                    var row = new XElement("row", new XAttribute("r", r), new XAttribute("spans", $"{1}:{width}"), new XAttribute(x14ac + "dyDescent", "0.25"));

                    // Iterate the cells
                    for (int c = 0; c < Rows[i].Count; c++)
                    {
                        int col = c + 1;

                        // Strings, numbers, and booleans are different from other datatypes
                        if (Rows[i][c].DataType == DataType.String)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "s"), new XElement("v", SharedStrings[Rows[i][c].Value])));
                        }
                        else if (Rows[i][c].DataType == DataType.Number)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XElement("v", Rows[i][c].Value)));
                        }
                        if (Rows[i][c].DataType == DataType.Boolean)
                        {
                            row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("t", "b"), new XElement("v", Rows[i][c].Value)));
                        }
                        else
                        {
                            for (int n = 0; n < CellFormats.Count; n++)
                            {
                                if (CellFormats[n].type == (int)Rows[i][c].DataType)
                                {
                                    row.Add(new XElement("c", new XAttribute("r", $"{Utilities.GetIndex(col)}{r.ToString()}"), new XAttribute("s", CellFormats[n].index.ToString()), new XElement("v", Rows[i][c].Value)));
                                }
                            }
                        }
                    }

                    sheetData.Add(row);
                }
#endif

            }

            return sheetData;
        }

        private int GetColumnHeaderCount()
        {
            int width = 0;
            foreach (Row dc in Rows)
            {
                width = dc.Count;
                break;
            }
            return width;
        }

        private int GetColumnCount()
        {
            int count = 0;
            foreach (Row dc in Rows)
            {
                count = dc.Count;
                break;
            }
            return count;
        }
    }
}
