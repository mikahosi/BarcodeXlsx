using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

using ClosedXML;
using ClosedXML.Utils;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace BarcodeXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            DecodeArgumentParamaters param = new DecodeArgumentParamaters(args);

            int imageNumber = 0;
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

                            BarcodeLib.Barcode barcode = new BarcodeLib.Barcode();
                            barcode.Height = 64;
                            barcode.Width = 256;
                            barcode.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                            barcode.BackColor = Color.White;
                            barcode.ImageFormat = ImageFormat.Bmp;
                            barcode.LabelPosition = BarcodeLib.LabelPositions.BOTTOMCENTER;
                            barcode.IncludeLabel = true;
                            barcode.Encode(DecodeBarcodeStyle(barcodeType), barcodeValue);

                            MemoryStream tempStream = new MemoryStream();
                            barcode.EncodedImage.Save(tempStream, ImageFormat.Png);
                            var picture = sheet.AddPicture(tempStream);
                            picture.MoveTo(cell);
                            picture.Scale(0.5, true);
                            picture.Height = (int )(cell.WorksheetRow().Height / 0.75);
                            picture.Width = (int )(cell.WorksheetColumn().Width / 0.118);
                        }
                    }
                }
            }

            book.Save();
        }

        static BarcodeLib.TYPE DecodeBarcodeStyle(string barcodeStyle)
        {
            BarcodeLib.TYPE barcodeType = BarcodeLib.TYPE.CODE128;
            switch (barcodeStyle.ToUpper())
            {
                case "JAN13":
                    barcodeType = BarcodeLib.TYPE.JAN13;
                    break;

                case "EAN13":
                    barcodeType = BarcodeLib.TYPE.EAN13;
                    break;

                case "EAN8":
                    barcodeType = BarcodeLib.TYPE.EAN8;
                    break;

                case "CODE128":
                    barcodeType = BarcodeLib.TYPE.CODE128;
                    break;
            }

            return barcodeType;
        }
    }
}
