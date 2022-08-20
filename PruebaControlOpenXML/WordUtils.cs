using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace PruebaControlOpenXML
{
    public class WordUtils
    {
        public static void SetPageSize(SectionProperties secProps, PageSizeTypes pageSizeTypes, PageOrientationValues pageOrientation = PageOrientationValues.Portrait)
        {
            PageSize pgSz = secProps.Descendants<PageSize>().FirstOrDefault();

            if (pgSz == null)
            {
                pgSz = new PageSize();
                secProps.InsertAt(pgSz, 0);
            }


            // Size in twentieths of a point
            var width = 0.0;
            var height = 0.0;
            switch (pageSizeTypes)
            {
                case PageSizeTypes.A4:
                    width = 11906.0;
                    height = 1684.0;
                    break;
                    
                case PageSizeTypes.A5:
                    width = 7096.0;
                    height = 0;
                    break;
                case PageSizeTypes.A6:
                    break;
                case PageSizeTypes.Letter:
                    break;
                case PageSizeTypes.Legal:
                    break;
                case PageSizeTypes.Executive:
                    break;
                case PageSizeTypes.Ledger:
                    break;
                case PageSizeTypes.Tabloid:
                    break;
                case PageSizeTypes.Custom:
                    break;
            }
        }
    }
}
