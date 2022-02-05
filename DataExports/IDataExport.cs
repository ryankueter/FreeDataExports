/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using FreeDataExports.Delimited;
using FreeDataExports.Spreadsheets.XL2019;
using FreeDataExports.Spreadsheets.XL2021;
using FreeDataExports.Spreadsheets.Ods1_3;

namespace FreeDataExports
{
    public interface IDataExport
    {
        XLSX2021 CreateXLSX2021();
        XLSX2019 CreateXLSX2019();
        ODSv1_3 CreateODSv1_3();
        Csv CreateCsv();
    }
}
