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
            try
            {
                Console.WriteLine("BarcodeXlsx v0.1");

                DecodeArgumentParamaters param = new DecodeArgumentParamaters(args);

                if (param.enabledVerbose)
                {
                    Console.WriteLine("--source {0}", param.sourceFileName);
                }

                XLWorkbook book = new XLWorkbook(param.sourceFileName);
                foreach (var sheet in book.Worksheets)
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
                                if (delimiterPos > 0)
                                {
                                    string barcodeType = barcodeData.Substring(0, delimiterPos);
                                    string barcodeValue = barcodeData.Substring(delimiterPos + 1);

                                    if (param.enabledVerbose)
                                    {
                                        Console.WriteLine("sheetName = {0}, cellAddress = {1}, barcodeType = {2}, barcodeValue = {3}", sheet.Name, cell.Address, barcodeType, barcodeValue);
                                    }

                                    try
                                    {
                                        BarcodeLib.Barcode barcode = new BarcodeLib.Barcode();
                                        barcode.Height = param.imageHeight;
                                        barcode.Width = param.imageWidht;
                                        barcode.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                                        barcode.IncludeLabel = param.enabledLabel;
                                        barcode.LabelPosition = BarcodeLib.LabelPositions.BOTTOMCENTER;
                                        barcode.LabelFont = new Font(FontFamily.GenericSansSerif, 8);
                                        barcode.BackColor = Color.White;
                                        barcode.ImageFormat = ImageFormat.Bmp;
                                        barcode.Encode(DecodeBarcodeStyle(barcodeType), barcodeValue);

                                        MemoryStream tempStream = new MemoryStream();
                                        barcode.EncodedImage.Save(tempStream, ImageFormat.Png);
                                        var picture = sheet.AddPicture(tempStream);
                                        picture.MoveTo(cell);
                                        picture.Scale(0.5, true);
                                        picture.Height = (int)(cell.WorksheetRow().Height / 0.75);
                                        picture.Width = (int)(cell.WorksheetColumn().Width / 0.118);
                                    }
                                    catch (Exception exp)
                                    {
                                        Console.Error.WriteLine("{0}", exp.Message);
                                    }
                                }
                            }
                        }
                    }
                }

                if (param.destinationFileName == null)
                {
                    book.Save();
                }
                else
                {
                    book.SaveAs(param.destinationFileName);
                }
            }
            catch (Exception exp)
            {
                Console.Error.WriteLine("{0}", exp.Message);
            }
        }

        /// <summary>
        /// Decode barcode style name to enum value
        /// </summary>
        /// <param name="barcodeStyle"></param>
        /// <returns></returns>
        static BarcodeLib.TYPE DecodeBarcodeStyle(string barcodeStyle)
        {
            BarcodeLib.TYPE barcodeType = BarcodeLib.TYPE.CODE128;
            switch (barcodeStyle.ToUpper())
            {
                case "UNSPECIFIED":
                    barcodeType = BarcodeLib.TYPE.UNSPECIFIED;
                    break;

                case "UPCA":
                    barcodeType = BarcodeLib.TYPE.UPCA;
                    break;

                case "UPCE":
                    barcodeType = BarcodeLib.TYPE.UPCE;
                    break;

                case "UPC_SUPPLEMENTAL_2DIGIT":
                    barcodeType = BarcodeLib.TYPE.UPC_SUPPLEMENTAL_2DIGIT;
                    break;

                case "UPC_SUPPLEMENTAL_5DIGIT":
                    barcodeType = BarcodeLib.TYPE.UPC_SUPPLEMENTAL_5DIGIT;
                    break;

                case "EAN13":
                    barcodeType = BarcodeLib.TYPE.EAN13;
                    break;

                case "EAN8":
                    barcodeType = BarcodeLib.TYPE.EAN8;
                    break;

                case "Interleaved2of5":
                    barcodeType = BarcodeLib.TYPE.Interleaved2of5;
                    break;

                case "Interleaved2of5_Mod10":
                    barcodeType = BarcodeLib.TYPE.Interleaved2of5_Mod10;
                    break;

                case "Standard2of5":
                    barcodeType = BarcodeLib.TYPE.Standard2of5;
                    break;

                case "Standard2of5_Mod10":
                    barcodeType = BarcodeLib.TYPE.Standard2of5_Mod10;
                    break;

                case "Industrial2of5":
                    barcodeType = BarcodeLib.TYPE.Industrial2of5;
                    break;

                case "Industrial2of5_Mod10":
                    barcodeType = BarcodeLib.TYPE.Industrial2of5_Mod10;
                    break;

                case "CODE39":
                    barcodeType = BarcodeLib.TYPE.CODE39;
                    break;

                case "CODE39Extended":
                    barcodeType = BarcodeLib.TYPE.CODE39Extended;
                    break;

                case "CODE39_Mod43":
                    barcodeType = BarcodeLib.TYPE.CODE39_Mod43;
                    break;

                case "Codabar":
                    barcodeType = BarcodeLib.TYPE.Codabar;
                    break;

                case "PostNet":
                    barcodeType = BarcodeLib.TYPE.PostNet;
                    break;

                case "BOOKLAND":
                    barcodeType = BarcodeLib.TYPE.BOOKLAND;
                    break;

                case "ISBN":
                    barcodeType = BarcodeLib.TYPE.ISBN;
                    break;

                case "JAN13":
                    barcodeType = BarcodeLib.TYPE.JAN13;
                    break;

                case "MSI_Mod10":
                    barcodeType = BarcodeLib.TYPE.MSI_Mod10;
                    break;

                case "MSI_2Mod10":
                    barcodeType = BarcodeLib.TYPE.MSI_2Mod10;
                    break;

                case "MSI_Mod11":
                    barcodeType = BarcodeLib.TYPE.MSI_Mod11;
                    break;

                case "MSI_Mod11_Mod10":
                    barcodeType = BarcodeLib.TYPE.MSI_Mod11_Mod10;
                    break;

                case "Modified_Plessey":
                    barcodeType = BarcodeLib.TYPE.Modified_Plessey;
                    break;

                case "CODE11":
                    barcodeType = BarcodeLib.TYPE.CODE11;
                    break;

                case "USD8":
                    barcodeType = BarcodeLib.TYPE.USD8;
                    break;

                case "UCC12":
                    barcodeType = BarcodeLib.TYPE.UCC12;
                    break;

                case "UCC13":
                    barcodeType = BarcodeLib.TYPE.UCC13;
                    break;

                case "LOGMARS":
                    barcodeType = BarcodeLib.TYPE.LOGMARS;
                    break;

                case "CODE128":
                    barcodeType = BarcodeLib.TYPE.CODE128;
                    break;

                case "CODE128A":
                    barcodeType = BarcodeLib.TYPE.CODE128A;
                    break;

                case "CODE128B":
                    barcodeType = BarcodeLib.TYPE.CODE128B;
                    break;

                case "CODE128C":
                    barcodeType = BarcodeLib.TYPE.CODE128C;
                    break;

                case "ITF14":
                    barcodeType = BarcodeLib.TYPE.ITF14;
                    break;

                case "CODE93":
                    barcodeType = BarcodeLib.TYPE.CODE93;
                    break;

                case "TELEPEN":
                    barcodeType = BarcodeLib.TYPE.TELEPEN;
                    break;

                case "FIM":
                    barcodeType = BarcodeLib.TYPE.FIM;
                    break;

                case "PHARMACODE":
                    barcodeType = BarcodeLib.TYPE.PHARMACODE;
                    break;

            }

            return barcodeType;
        }
    }
}
