/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
namespace FreeDataExports
{
    /// <summary>
    /// The list of available datatypes
    /// </summary>
    public enum DataType : int
    {
        String = 0,
        Number = 1,
        Currency = 2,
        Percent = 3,
        DateTime = 4,
        Decimal = 5,
        Date = 6,
        Time = 7,
        LongDate = 8,
        DateTime24 = 9,
        Time24 = 10,
        Boolean = 11
    }
}
