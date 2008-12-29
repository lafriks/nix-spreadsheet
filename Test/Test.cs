using System;
using System.Collections;
using System.Data;
using System.Text;
using System.Security.Cryptography;

using Nix.SpreadSheet;
using Nix.Zip;

namespace Test
{
    /// <summary>
    /// Summary description for Test.
    /// </summary>
    public class Test
    {
        // Generic function using HashAlgorithm 
        public static void HashData(HashAlgorithm hashAlg, string str )
        {
            byte[] rawBytes = System.Text.ASCIIEncoding.ASCII.GetBytes( str );
            byte[] hashData = hashAlg.ComputeHash( rawBytes );
            Console.WriteLine( BitConverter.ToString( hashData ) );
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            /*Adler32 ad32 = new Adler32();
            HashData(ad32, "Wikipedia");
            CRC32 crc32 = new CRC32(CRC32.DefaultPolynomialLE);
            HashData(crc32, "Wikipedia");*/
            SpreadSheetDocument doc = new SpreadSheetDocument();
            Sheet s = doc.AddSheet();
            //s["IV65536"].Value = "MAX";
            doc.Save(@"C:\test.xls", SpreadSheetFileFormat.XlsBinary);
        }
    }
}
