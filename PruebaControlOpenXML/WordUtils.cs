using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;

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
                pgSz.Width = (UInt32Value)ConvertInchToTwip(width);
                pgSz.Height = (UInt32Value)ConvertInchToTwip(height);
            }
            else
            {
                pgSz.Width = (UInt32Value)ConvertInchToTwip(height);
                pgSz.Height = (UInt32Value)ConvertInchToTwip(width);
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
                pgMar.Top = (Int32Value)ConvertInchToTwip(margin.top);
                pgMar.Bottom = (Int32Value)ConvertInchToTwip(margin.bottom);
                pgMar.Left = (UInt32Value)ConvertInchToTwip(margin.left);
                pgMar.Right = (UInt32Value)ConvertInchToTwip(margin.right);
            }
            else
            {
                pgMar.Top = (Int32Value)ConvertInchToTwip(margin.left);
                pgMar.Bottom = (Int32Value)ConvertInchToTwip(margin.right);
                pgMar.Left = (UInt32Value)ConvertInchToTwip(margin.bottom);
                pgMar.Right = (UInt32Value)ConvertInchToTwip(margin.top);
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
                    width = 8.27;
                    height = 11.69;
                    break;

                case PageSizeTypes.A5:
                    width = 5.83;
                    height = 8.27;
                    break;
                case PageSizeTypes.A6:
                    width = 5.83;
                    height = 8.27;
                    break;
                case PageSizeTypes.Custom:
                    break;
            }
            
            return (width, height);
        }
        #endregion


        #region Comandos para personalizar textos en tablas
        public static string SetBold()
        {
            return "[N]";
        }

        public static string SetItalic()
        {
            return "[I]";
        }

        public static string SetUnderline()
        {
            return "[U]";
        }

        public static string SetFontSize(int size)
        {
            return $"[F:{size}]";
        }

        public static string SetFontColor(string color)
        {
            return $"[FC:{color}]";
        }

        public static string SetCellColor(string color)
        {
            return $"[CC:{color}]";
        }

        public static string SetLeftAligment()
        {
            return "¬";
        }
        #endregion


        #region Conversor de unidades
        /// <summary>
        /// Convertir pulgada a twip
        /// </summary>
        /// <param name="valcm"></param>
        /// <returns></returns>
        public static double ConvertInchToTwip(double valp)
        {
            return valp * 1440;
        }

        /// <summary>
        /// Convertir cm a twip
        /// </summary>
        /// <param name="valcm"></param>
        /// <returns></returns>
        public static double ConvertCmToTwip(double valcm)
        {
            return valcm * 567;
        }

        /// <summary>
        /// Convertir twip a cm
        /// </summary>
        /// <param name="valtwip"></param>
        /// <returns></returns>
        public static double ConvertTwipToCm(double valtwip)
        {
            return valtwip * (1 / 567);
        }

        /// <summary>
        /// Convertir twip a cm
        /// </summary>
        /// <param name="valtwip"></param>
        /// <returns></returns>
        public static double ConvertTwipToInch(double valtwip)
        {
            return valtwip * (1 / 1440);
        }
        #endregion


        public static void SaveDocumentAsPdf(string docFullPath, string pdfFullPath)
        {
            Application appWord = new Application();
            appWord.Visible = false;
            var wordDocument = appWord.Documents.Open(docFullPath);

            wordDocument.SaveAs2(pdfFullPath, WdSaveFormat.wdFormatPDF);
            wordDocument.Close(false);
            appWord.Quit();
        }
    }
}
