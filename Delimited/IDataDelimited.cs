/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System.Threading.Tasks;

namespace FreeDataExports.Delimited
{
    public interface IDataDelimited
    {
        void AddRow(params object[] o);
        Task SaveAsync(string path);
        void Save(string path);
        Task<byte[]> GetBytesAsync();
        byte[] GetBytes();
        string GetErrors();
    }
}
