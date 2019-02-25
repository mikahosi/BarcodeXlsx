using System;
using System.Text;
using System.Text.RegularExpressions;

using ClosedXML;
using ClosedXML.Excel;

namespace BarcodeXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            DecodeArgumentParamaters param = new DecodeArgumentParamaters(args);

            XLWorkbook book = new XLWorkbook(param.sourceFileName);
            foreach(var sheet in book.Worksheets)
            {
                foreach (var cell in sheet.Cells())
                {
                    string cellValue = cell.GetString();
                    if (cellValue.Length > param.preChars.Length + param.postChars.Length)
                    {
                        string preChars = cellValue.Substring(0, param.preChars.Length);
                        string postChears = cellValue.Substring(cellValue.Length - param.postChars.Length);

                        if (preChars == param.preChars && postChears == param.postChars)
                        {
                            string barcodeData = cellValue.Substring(param.preChars.Length, cellValue.Length - param.preChars.Length - param.postChars.Length);
                            int delimiterPos = barcodeData.IndexOf(":");
                            string barcodeType = barcodeData.Substring(0, delimiterPos);
                            string barcodeValue = barcodeData.Substring(delimiterPos + 1);
                            Console.WriteLine("barcodeType = {0}, barcodeValue = {1}", barcodeType, barcodeValue);
                        }
                    }
                }
            }
        }
    }
}
