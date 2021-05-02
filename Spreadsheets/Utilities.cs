/**
 * Author: Ryan A. Kueter
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
using System;

namespace FreeDataExports
{
    internal static class Utilities
    {
        /// <summary>
        /// A Guid for xlsx files
        /// </summary>
        /// <returns></returns>
        internal static string NewGuid()
        {
            return Guid.NewGuid().ToString("B").ToUpper();
        }

        /// <summary>
        /// A timestamp for xlsx files
        /// </summary>
        /// <returns></returns>
        internal static string XlsxTimestamp()
        {
            return DateTime.Now.ToString("yyyy-MM-ddTHH\\:mm\\:ssZ");
        }

        /// <summary>
        /// A timestamp for the ODS files
        /// </summary>
        /// <returns></returns>
        internal static string OdsTimestamp()
        {
            return DateTime.Now.ToString("yyyy-MM-ddTHH\\:mm\\:ss.fffffff");
        }

        /// <summary>
        /// Returns an alpha index for a column
        /// e.g., A2
        /// </summary>
        /// <param name="num">The numeric index for the column</param>
        /// <returns></returns>
        internal static string GetIndex(int num)
        {
            string str = String.Empty;
            char achar;
            int mod;
            num--;
            while (true)
            {
                mod = (num % 26) + 65;
                num = (int)(num / 26);
                achar = (char)mod;

                str = achar + str;
                if (num > 0)
                {
                    num--;
                }
                else if (num == 0)
                {
                    break;
                }
            }
            return str;
        }
    }
}
