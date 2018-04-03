using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;

namespace PresentationToPDF {
    /// <summary>
    /// Contains basic information about a PowerPoint presentation; Path, Size, Slides, etc.
    /// </summary>
    class PptInfo : OfficeFileInfo {
        /// <summary>
        /// Instantiates a new instance of the <code>PresentationToPdf.PptInfo</code> class.
        /// </summary>
        /// <param name="filePath">Path to the PowerPoint presentation</param>
        /// <exception cref="System.IO.PathTooLongException"></exception>
        /// <exception cref="System.IO.UnauthorizedAccessException"></exception>
        public PptInfo(string filePath)
            : base(filePath) {
        }

        public override string PageString {
            get {
                if (Pages >= 0) {
                    return String.Format("| {0} Slide(s)", Pages);
                }
                else {
                    return string.Empty;
                }
            }

            protected set { } // do nothing
        }

        protected override int CountPages() {
            int count = 0;
            bool compatMode = System.IO.Path.GetExtension(Path) != ".pptx";

            if (compatMode) { // use PP to get slide count (for older ppt files)
                // TODO - improve slide count speed for ppt files (slide count disable for now)
                //var ppApp = new Application();
                //Presentation p =  ppApp.Presentations.Open(Path, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                //count = p.Slides.Count;
                //p.Close();
                //ppApp.Quit();
                return -1;
            }
            else { // use OpenXML method
                using (var doc = PresentationDocument.Open(Path, false)) {
                    PresentationPart pres = doc.PresentationPart;

                    if (pres != null && pres.SlideParts != null) {
                        count = pres.SlideParts.Count();
                    }
                }
            }

            return count;
        }
    }
}
