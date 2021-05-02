/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;

namespace FreeDataExports.Spreadsheets.Ods1_3
{
    /// <summary>
    /// A class to store each cell value.
    /// This can also be used to handle data type
    /// errors. 
    /// </summary>
    internal class Cell : IDataCell
    {
        public object Value { get; set; } // The raw value
        public string FormattedValue { get; set; } // The final value
        public DataType DataType { get; set; } // The datatype specified
        public string Errors { get; set; } // Stores conversion errors

        /// <summary>
        /// The constructor for the cell.
        /// </summary>
        /// <param name="value">The value to store in the cell.</param>
        /// <param name="data">The value's datatype.</param>
        public Cell(object value, DataType data)
        {
            DataType = data;
            Value = GetValue(value);
        }

        /// <summary>
        /// Conversion errors can be caught here.
        /// 
        /// Some potential formating options:
        /// using System.Globalization;
        /// 
        /// Currency:
        /// return Decimal.Parse(value.ToString(), NumberStyles.Currency).ToString();
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private object GetValue(object value)
        {
            // Handle null values
            if (value == null)
            {
                FormattedValue = String.Empty;
                return null;
            }

            try
            {
                switch (DataType)
                {
                    case DataType.String:
                        return value.ToString().Trim();
                    case DataType.Number:
                        FormattedValue = value.ToString();
                        return Decimal.Parse(value.ToString());
                    case DataType.DateTime:
                        var dt = Convert.ToDateTime(value);
                        FormattedValue = dt.ToString("MM/dd/yyyy hh:mm tt");
                        return dt.ToString("yyyy-MM-ddTHH\\:mm\\:ss");
                    case DataType.Currency:
                        FormattedValue = value.ToString();
                        return Decimal.Parse(value.ToString());
                    case DataType.Percent:
                        FormattedValue = value.ToString();
                        return Decimal.Parse(value.ToString());
                    case DataType.Decimal:
                        FormattedValue = value.ToString();
                        return Decimal.Parse(value.ToString());
                    case DataType.Date:
                        var d = Convert.ToDateTime(value);
                        FormattedValue = d.ToString("MM/dd/yyyy");
                        return d.ToString("yyyy-MM-dd");
                    case DataType.Time:
                        var t = Convert.ToDateTime(value);
                        FormattedValue = t.ToString("HH:mm:ss tt");
                        return t.ToString("PThh\\Hmm\\Mss\\S");
                    case DataType.LongDate:
                        var ld = Convert.ToDateTime(value);
                        FormattedValue = ld.ToString("MMMM dd, yyyy");
                        return ld.ToString("yyyy-MM-dd");
                    case DataType.DateTime24:
                        var dt24 = Convert.ToDateTime(value);
                        FormattedValue = dt24.ToString("MM/dd/yyyy HH:mm:ss");
                        return dt24.ToString("yyyy-MM-ddTHH\\:mm\\:ss");
                    case DataType.Time24:
                        var t24 = Convert.ToDateTime(value);
                        FormattedValue = t24.ToString("HH:mm:ss");
                        return t24.ToString("PTHH\\Hmm\\Mss\\S");
                    case DataType.Boolean:
                        FormattedValue = value.ToString().ToUpper();
                        return value.ToString().ToLower();
                    default:
                        FormattedValue = String.Empty;
                        return null;
                }
            }
            catch (Exception ex)
            {
                Errors = ex.Message;
                return null;
            }
        }
    }
}
