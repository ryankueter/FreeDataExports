/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using FreeDataExports.Delimited;
using FreeDataExports.Spreadsheets.XL2019;
using FreeDataExports.Spreadsheets.Ods1_3;

namespace FreeDataExports
{
    public class DataExport : IDataExport
    {
        public XLSX2019 CreateXLSX2019()
        {
            return new XLSX2019();
        }
        public ODSv1_3 CreateODSv1_3()
        {
            return new ODSv1_3();
        }
        public Csv CreateCsv()
        {
            return new Csv();
        }
    }
}
