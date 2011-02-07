using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Nix.SpreadSheet.Provider
{
    /// <summary>
    /// CSV file format provider 
    /// </summary>
    public class CsvFileFormatProvicer : IFileFormatProvider
    {
        public void Save(SpreadSheetDocument document, System.IO.Stream stream)
        {
            foreach (Sheet s in document)
            {
                TextWriter t = new StreamWriter(stream, Encoding.UTF8);
                for (int r = 0; r <= s.LastRow; r++ )
                {
                    for (int c = 0; c <= s.LastColumn; c++)
                    {
                        if (c != 0)
                            t.Write(",");
                        object val = s[r][c].Value;
                        if (val == null || (val is string && (string)val == string.Empty))
                            t.Write("\"\"");
			            else if (val is int || val is float || val is double ||
                	            val is decimal || val is long || val is short)
            	            t.Write(val);
			            else if (val is ErrorCode)
                            t.Write("\"" + ((ErrorCode)val).ToString("g") + "\"");
                        else
                            t.Write("\"" + val.ToString().Replace("\"", "\"\"") + "\"");
                    }
                    t.WriteLine();
                }
                break;
            }
        }
    }
}
