/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
namespace FreeDataExports.Spreadsheets
{
    public interface IDataCell
    {
        object Value { get; set; }
        string FormattedValue { get; set; }
        DataType DataType { get; set; }
        string Errors { get; set; }
    }
}
