using System;
using System.Collections.Generic;
using System.Text;

namespace BarcodeXlsx
{
    class DecodeArgumentParamaters
    {
        public string preChars;
        public string postChars;
        public string sourceFileName;
        public string destinationFileName;
        public bool enabledVerbose = false;
        public bool enabledProgress = false;
        public bool enabledLabel = false;
        public int imageWidht = 256;
        public int imageHeight = 64;

        public DecodeArgumentParamaters(string[] args)
        {
            preChars = "{";
            postChars = "}";
            enabledProgress = false;
            enabledVerbose = false;

            bool sourceFlag = false;
            bool destinationFlag = false;
            bool preCharsFlag = false;
            bool postCharsFlag = false;
            bool imageWidthFlag = false;
            bool imageHeightFlag = false;

            foreach (var arg in args)
            {
                if (arg == "-source")
                {
                    sourceFlag = true;
                }
                else if (arg == "-destination")
                {
                    destinationFlag = true;
                }
                else if (arg == "-prechar")
                {
                    preCharsFlag = true;
                }
                else if (arg == "-postchar")
                {
                    postCharsFlag = true;
                }
                else if (arg == "-width")
                {
                    imageWidthFlag = true;
                }
                else if (arg == "-height")
                {
                    imageHeightFlag = true;
                }
                else if (arg == "-showlabel")
                {
                    enabledLabel = true;
                }
                else if (arg == "-progress")
                {
                    enabledProgress = true;
                }
                else if (arg == "-verbose")
                {
                    enabledVerbose = true;
                }
                else if (sourceFlag)
                {
                    sourceFlag = false;
                    sourceFileName = arg;
                }
                else if (destinationFlag)
                {
                    destinationFlag = false;
                    destinationFileName = arg;
                }
                else if (preCharsFlag)
                {
                    preCharsFlag = false;
                    preChars = arg;
                }
                else if (postCharsFlag)
                {
                    postCharsFlag = false;
                    postChars = arg;
                }
                else if (imageWidthFlag)
                {
                    imageWidthFlag = false;
                    if (int.TryParse(arg, out imageWidht))
                    {

                    }
                }
                else if (imageHeightFlag)
                {
                    imageHeightFlag = false;
                    if (int.TryParse(arg, out imageHeight))
                    {

                    }
                }
            }
        }
    }
}
