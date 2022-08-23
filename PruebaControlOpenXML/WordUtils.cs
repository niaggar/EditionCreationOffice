﻿using System;
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
        /// <summary>
        /// Establece el tamano de pagina para una seccion
        /// </summary>
        /// <param name="secProps">props de la seccion</param>
        /// <param name="pageSizeTypes">tipo de hoja</param>
        /// <param name="pageOrientation">orientacion de la hoja</param>
        public static void SetPageSize(SectionProperties secProps, PageSizeTypes pageSizeTypes, PageOrientationValues pageOrientation = PageOrientationValues.Portrait)
        {
            PageSize pgSz = secProps.Descendants<PageSize>().FirstOrDefault();

            if (pgSz == null)
            {
                pgSz = new PageSize();
                secProps.InsertAt(pgSz, 0);
            }

            var (width, height) = GetPaperSize(pageSizeTypes);

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

        /// <summary>
        /// Establece el tamano de margen de una seccion
        /// </summary>
        /// <param name="secProps">props de la seccion</param>
        /// <param name="margin">margenes, dados en twips</param>
        /// <param name="pageOrientation">orientacion de la pagina</param>
        public static void SetMarginSize(SectionProperties secProps, (double top, double right, double bottom, double left) margin, PageOrientationValues pageOrientation = PageOrientationValues.Portrait)
        {
            var pgMar = secProps.Descendants<PageMargin>().FirstOrDefault();
            if (pgMar == null)
            {
                pgMar = new PageMargin();
                secProps.InsertAt(pgMar, 0);
            }

            if (pageOrientation == PageOrientationValues.Portrait)
            {
                pgMar.Top = (Int32Value)margin.top;
                pgMar.Bottom = (Int32Value)margin.bottom;
                pgMar.Left = (UInt32Value)margin.left;
                pgMar.Right = (UInt32Value)margin.right;
            }
            else
            {
                pgMar.Top = (Int32Value)margin.left;
                pgMar.Bottom = (Int32Value)margin.right;
                pgMar.Left = (UInt32Value)margin.bottom;
                pgMar.Right = (UInt32Value)margin.top;
            }
        }

        /// <summary>
        /// Obtener tamano de pagina de una seccion dada
        /// </summary>
        /// <param name="secProps"></param>
        /// <returns></returns>
        public static (double width, double height) GetPageSize(SectionProperties secProps)
        {
            PageSize pgSz = secProps.Descendants<PageSize>().FirstOrDefault();
            if (pgSz == null)
            {
                return (0, 0);
            }
            else
            {
                return (pgSz.Width.Value, pgSz.Height.Value);
            }
        }

        /// <summary>
        /// Obtener margenes de una seccion dada
        /// </summary>
        /// <param name="secProps"></param>
        /// <returns></returns>
        public static (double top, double right, double bottom, double left) GetMarginSize(SectionProperties secProps)
        {
            PageSize pgSz = secProps.Descendants<PageSize>().FirstOrDefault();
            PageMargin pgMar = secProps.Descendants<PageMargin>().FirstOrDefault();

            if (pgMar == null) return (0, 0, 0, 0);

            if (pgSz.Orient == PageOrientationValues.Portrait)
                return (pgMar.Top, pgMar.Right, pgMar.Bottom, pgMar.Left);
            else
                return (pgMar.Left, pgMar.Top, pgMar.Right, pgMar.Bottom);
        }

        /// <summary>
        /// Obtener el tamano de pagina para un tipo dado
        /// </summary>
        /// <param name="pageSizeTypes"></param>
        /// <returns></returns>
        public static (double width, double height) GetPaperSize(PageSizeTypes pageSizeTypes)
        {
            var width = 0.0;
            var height = 0.0;

            // Los valores estan dados twips = 0.0017638889 cm (o en 0.05 puntos, donde 72 puntos = 1 pulgada)
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
            
            return (width, height);
        }
        #endregion


        #region Conversor de unidades
        /// <summary>
        /// Convertir cm a twip
        /// </summary>
        /// <param name="valcm"></param>
        /// <returns></returns>
        public static double ConvertCmToTwip(double valcm)
        {
            return valcm / 0.0017638889;
        }

        /// <summary>
        /// Convertir twip a cm
        /// </summary>
        /// <param name="valtwip"></param>
        /// <returns></returns>
        public static double ConvertTwipToCm(double valtwip)
        {
            return valtwip * 0.0017638889;
        }
        #endregion
    }
}
