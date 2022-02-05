/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;

namespace FreeDataExports.Spreadsheets.XL2021
{
    internal class Cell : IDataCell
    {
        public object Value { get; set; }
        public string FormattedValue { get; set; }
        public DataType DataType { get; set; }
        public string Errors { get; set; }
        public Cell(object value, DataType data)
        {
            DataType = data;
            Value = GetValue(value);
        }

        /// <summary>
        /// If this triggers a conversion error,
        /// they can be caught here.
        /// 
        /// Some potential formating options:
        /// using System.Globalization;
        /// 
        /// Currency:
        /// return Decimal.Parse(value.ToString(), NumberStyles.Currency).ToString();
        /// 
        /// From OA Date:
        /// DateTime.FromOADate((float)o)
        /// 
        /// OA Time:
        /// var d = DateTime.Parse(o.ToString());
        /// var date = (float)d.ToOADate();
        /// return date - Math.Truncate(date);
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private object GetValue(object value)
        {
            // Handle null values
            if (value == null)
            {
                return null;
            }

            try
            {
                switch (DataType)
                {
                    case DataType.String:
                        return value.ToString().Trim();
                    case DataType.Number:
                        return Decimal.Parse(value.ToString());
                    case DataType.DateTime:
                        return ConvertToOADate(value);
                    case DataType.Currency:
                        return Decimal.Parse(value.ToString());
                    case DataType.Percent:
                        return Decimal.Parse(value.ToString());
                    case DataType.Decimal:
                        return Decimal.Parse(value.ToString());
                    case DataType.Date:
                        return ConvertToOADate(value);
                    case DataType.Time:
                        return ConvertToOADate(value);
                    case DataType.LongDate:
                        return ConvertToOADate(value);
                    case DataType.DateTime24:
                        return ConvertToOADate(value);
                    case DataType.Time24:
                        return ConvertToOADate(value);
                    case DataType.Boolean:
                        int b = 0;
                        switch (value)
                        {
                            case true:
                                b = 1;
                                break;
                            case false:
                                b = 0;
                                break;
                            case 1:
                                b = 1;
                                break;
                            case 0:
                                b = 0;
                                break;
                        }
                        return b;
                    default:
                        return value;
                }
            }
            catch (Exception ex)
            {
                Errors = ex.Message;
                return null;
            }
        }

        private object ConvertToOADate(object o)
        {
            var d = DateTime.Parse(o.ToString());
            return Math.Round(d.ToOADate(), 15);
        }
    }
}
