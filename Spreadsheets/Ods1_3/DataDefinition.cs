/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System.Xml.Linq;

namespace FreeDataExports.Spreadsheets.Ods1_3
{
    /// <summary>
    /// Stores the data definition for a type.
    /// </summary>
    internal class DataDefinition
    {
        internal int Index { get; set; }
        internal string Name { get; set; }
        internal DataType DataType { get; set; }
        internal XElement Element { get; set; }
        internal string Worksheet { get; set; }
    }
}
