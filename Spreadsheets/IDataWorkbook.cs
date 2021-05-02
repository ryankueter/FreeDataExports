/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System.Threading.Tasks;

namespace FreeDataExports.Spreadsheets
{
    public interface IDataWorkbook
    {
        IDataWorksheet AddWorksheet(string Name);
        string CreatedBy { get; set; }
        decimal FontSize { get; set; }
        void AddErrorsWorksheet();
        Task SaveAsync(string path);
        void Save(string path);
        Task<byte[]> GetBytesAsync();
        byte[] GetBytes();
        string GetErrors();
        void Format(DataType type, string format);
    }
}
