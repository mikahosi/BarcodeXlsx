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
        public bool enabledProgress = false;

        public DecodeArgumentParamaters(string[] args)
        {
            bool sourceFlag = false;
            bool destinationFlag = false;
            bool preCharsFlag = false;
            bool postCharsFlag = false;

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
                else if (arg == "-progress")
                {
                    enabledProgress = true;
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
            }
        }
    }
}
