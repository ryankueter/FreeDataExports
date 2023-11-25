/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
/// <summary>
/// Provides the data definitions for the data types
/// </summary>
namespace FreeDataExports.Spreadsheets.XL2019
{
    internal sealed class DataDefinition
    {
        internal int index { get; set; }
        internal int type { get; set; }
        internal string numFmtId { get; set; }
        internal string formatCode { get; set; }
    }
}
