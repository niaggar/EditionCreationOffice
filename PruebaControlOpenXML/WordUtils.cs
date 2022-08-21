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
        #region Controlar tamano de pagina y margenes
        public static void SetPageSize(SectionProperties secProps, PageSizeTypes pageSizeTypes, PageOrientationValues pageOrientation = PageOrientationValues.Portrait)
        {
            PageSize pgSz = secProps.Descendants<PageSize>().FirstOrDefault();

            if (pgSz == null)
            {
                pgSz = new PageSize();
                secProps.InsertAt(pgSz, 0);
            }

            // Los valores estan dados en 0.05 puntos, donde 72 puntos = 1 pulgada
            var width = 0.0;
            var height = 0.0;
            switch (pageSizeTypes)
            {
                case PageSizeTypes.A4:
                    width = 11900;
                    height = 16840;
                    break;

                case PageSizeTypes.A5:
                    width = 8400;
                    height = 11900;
                    break;

                case PageSizeTypes.A6:
                    width = 5960;
                    height = 8400;
                    break;

                case PageSizeTypes.Custom:
                    break;
            }

            pgSz.Orient = new EnumValue<PageOrientationValues>(pageOrientation);
            if (pageOrientation == PageOrientationValues.Portrait)
            {
                pgSz.Width = (UInt32Value)width;
                pgSz.Height = (UInt32Value)height;
            }
            else
            {
                pgSz.Width = (UInt32Value)height;
                pgSz.Height = (UInt32Value)width;
            }
        }

        public static void SetMarginSize(SectionProperties secProps, double top, double bottom, double left, double right, PageOrientationValues pageOrientation = PageOrientationValues.Portrait)
        {
            var pgMar = secProps.Descendants<PageMargin>().FirstOrDefault();
            if (pgMar == null)
            {
                pgMar = new PageMargin();
                secProps.InsertAt(pgMar, 0);
            }

            if (pageOrientation == PageOrientationValues.Portrait)
            {
                pgMar.Top = (Int32Value)top;
                pgMar.Bottom = (Int32Value)bottom;
                pgMar.Left = (UInt32Value)left;
                pgMar.Right = (UInt32Value)right;
            }
            else
            {
                pgMar.Top = (Int32Value)left;
                pgMar.Bottom = (Int32Value)right;
                pgMar.Left = (UInt32Value)bottom;
                pgMar.Right = (UInt32Value)top;
            }
        }
        #endregion
    }
}
