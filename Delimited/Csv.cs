/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;
using System.IO;
using System.Threading.Tasks;
using System.Text;

namespace FreeDataExports.Delimited
{
    /// <summary>
    /// The Csv class
    /// </summary>
    public class Csv : IDataDelimited
    {
        public Csv()
        {
            Content = new StringBuilder();
            Errors = new StringBuilder();
        }
        // Stores the content of the file
        private StringBuilder Content { get; set; }
        private StringBuilder Errors { get; set; }
        /// <summary>
        /// A method that gets the data type conversion errors.
        /// </summary>
        /// <returns></returns>
        public string GetErrors()
        {
            return Errors.ToString();
        }

        /// <summary>
        /// A method for adding rows
        /// </summary>
        /// <param name="o">Cell data</param>
        public void AddRow(params object[] o)
        {
            var b = new StringBuilder();
            for (int i = 0; i < o.Length; i++)
            {
                try
                {
                    if (i == 0)
                    {
                        b.Append(FormatValue(o[i]));
                    }
                    else
                    {
                        b.Append(",");
                        b.Append(FormatValue(o[i]));
                    }
                }
                catch (Exception ex)
                {
                    Errors.AppendLine(ex.Message);
                }
            }
            Content.AppendLine(b.ToString());
        }

        /// <summary>
        /// An asynchronous save method
        /// </summary>
        /// <param name="path">Full path and file name with extension.</param>
        /// <returns></returns>
        public Task SaveAsync(string path)
        {
            return Task.Run(() => {
                var bytes = GetBytes();
                using (FileStream SourceStream = File.Open(path, FileMode.Create))
                {
                    SourceStream.Seek(0, SeekOrigin.End);
                    SourceStream.Write(bytes, 0, bytes.Length);
                }
            });
        }

        /// <summary>
        /// A synchronous save method
        /// </summary>
        /// <param name="path">Full path and file name with extension.</param>
        /// <returns></returns>
        public void Save(string path)
        {
            File.WriteAllText(path, Content.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// An asychronous GetBytes method
        /// </summary>
        /// <returns></returns>
        public Task<byte[]> GetBytesAsync()
        {
            return Task.Run(() => {
                return GetBytes();
            });
        }

        /// <summary>
        /// A sychronous GetBytes method
        /// </summary>
        /// <returns></returns>
        public byte[] GetBytes()
        {
            return Encoding.UTF8.GetBytes(Content.ToString());
        }

        /// <summary>
        /// Applies special formating
        /// RFC: https://tools.ietf.org/html/rfc4180
        /// </summary>
        /// <param name="o"></param>
        /// <returns></returns>
        private string FormatValue(object o)
        {
            // Convert the value to a string
            string value = o.ToString();

            // Escapes quotations.
            if (value.Contains("\"") == true)
            {
                value = value.Replace("\"", "\"\"");
            }

            // Inserts a quote before and after the string.
            if (value.Contains(",") == true)
            {
                value = value.Insert(0, "\"");
                value = value.Insert(value.Length, "\"");
            }

            return value;
        }
    }
}
