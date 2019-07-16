using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

using ClosedXML;
using ClosedXML.Utils;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using ImageMagick;

namespace Common
{
    public class BarcodeXlsxImporter
    {
        public string preChars = "{";
        public string postChars = "}";
        public bool enabledVerbose = false;
        public bool enabledProgress = false;
        public bool enabledLabel = false;
        public bool enabledRemoveTag = false;
        public int imageWidth = 256;
        public int imageHeight = 64;
        public int marginHeight = 4;
        public int marginWidth = 4;

        public BarcodeXlsxImporter()
        {

        }
        public void Convert(string sourceFileName)
        {
            XLWorkbook book = new XLWorkbook(sourceFileName);
            Convert(book);
            book.Save();
        }

        public void Convert(string sourceFileName, string destFileName)
        {
            XLWorkbook book = new XLWorkbook(sourceFileName);
            Convert(book);
            book.SaveAs(destFileName);
        }

        public void Convert(Stream inputStream, Stream outputStream)
        {
            XLWorkbook book = new XLWorkbook(inputStream);
            Convert(book);
            book.SaveAs(outputStream);
        }

        public void Convert(XLWorkbook book)
        {
            foreach (var sheet in book.Worksheets)
            {
                foreach (var cell in sheet.Cells())
                {
                    string cellValue = cell.GetString();
                    if (cellValue.Length > preChars.Length + postChars.Length)
                    {
                        string preChars = cellValue.Substring(0, this.preChars.Length);
                        string postChears = cellValue.Substring(cellValue.Length - this.postChars.Length);

                        if (preChars == this.preChars && postChears == this.postChars)
                        {
                            string barcodeData = cellValue.Substring(preChars.Length, cellValue.Length - preChars.Length - postChars.Length);
                            int delimiterPos = barcodeData.IndexOf(":");
                            if (delimiterPos > 0)
                            {
                                string barcodeType = barcodeData.Substring(0, delimiterPos);
                                string barcodeValue = barcodeData.Substring(delimiterPos + 1);

                                try
                                {
                                    BarcodeLib.Barcode barcode = new BarcodeLib.Barcode();
                                    barcode.Height = imageHeight - marginHeight * 2;
                                    barcode.Width = imageWidth - marginWidth * 2;
                                    barcode.Alignment = BarcodeLib.AlignmentPositions.CENTER;
                                    barcode.IncludeLabel = enabledLabel;
                                    barcode.LabelPosition = BarcodeLib.LabelPositions.BOTTOMCENTER;
                                    barcode.LabelFont = new Font(FontFamily.GenericSansSerif, 8);
                                    barcode.BackColor = Color.White;
                                    barcode.ImageFormat = ImageFormat.Bmp;
                                    barcode.Encode(DecodeBarcodeStyle(barcodeType), barcodeValue);

                                    MemoryStream tempStream1 = new MemoryStream();
                                    barcode.EncodedImage.Save(tempStream1, ImageFormat.Png);

                                    tempStream1.Position = 0;
                                    MagickImage image = new MagickImage(tempStream1);
                                    image.MatteColor = MagickColors.White;
                                    image.Frame(marginWidth, marginHeight, 0, 0);
                                    image.Transparent(MagickColors.White);
                                    MemoryStream tempStream2 = new MemoryStream();
                                    image.Write(tempStream2, MagickFormat.Png);

                                    var picture = sheet.AddPicture(tempStream2);
                                    picture.MoveTo(cell);
                                    picture.Scale(0.5, true);
                                    picture.Height = (int)(cell.WorksheetRow().Height / 0.75);
                                    picture.Width = (int)(cell.WorksheetColumn().Width / 0.118);

                                    if (enabledRemoveTag)
                                    {
                                        cell.SetValue("");
                                    }
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
        }

        /// <summary>
        /// Decode barcode style name to enum value
        /// </summary>
        /// <param name="barcodeStyle"></param>
        /// <returns></returns>
        public BarcodeLib.TYPE DecodeBarcodeStyle(string barcodeStyle)
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
