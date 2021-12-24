using FreeDataExports.Delimited;
using FreeDataExports.Spreadsheets.XL2019;
using FreeDataExports.Spreadsheets.Ods1_3;

namespace FreeDataExports
{
    public interface IDataExport
    {
        XLSX2019 CreateXLSX2019();
        ODSv1_3 CreateODSv1_3();
        Csv CreateCsv();
    }
}
