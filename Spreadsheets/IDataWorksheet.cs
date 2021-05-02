/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
namespace FreeDataExports.Spreadsheets
{
    public interface IDataWorksheet
    {
        string Name { get; }
        void AddRow(params IDataCell[] c);
        IDataCell AddCell(object Data, DataType Type);
        string TabColor { get; set; }
        void ColumnWidths(params string[] columnWidths);
    }
}
