using DIS.Reportes.Automatizados.Models.Enums;
using DIS.Reportes.Automatizados.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIS.Reportes.Automatizados.Components
{
    public class CreateSection
    {
        /// <summary>
        /// Crear nueva seccion
        /// </summary>
        /// <returns></returns>
        public static Paragraph NewS()
        {
            var paragraphSectionBreak = new Paragraph();
            var paragraphSectionBreakProperties = new ParagraphProperties();
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.NextPage });

            paragraphSectionBreakProperties.Append(SectionBreakProperties);
            paragraphSectionBreak.Append(paragraphSectionBreakProperties);

            return paragraphSectionBreak;
        }

        /// <summary>
        /// Crear seccion final del documento
        /// </summary>
        /// <returns></returns>
        public static SectionProperties NewSFinal()
        {
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.Continuous });

            return SectionBreakProperties;
        }

        /// <summary>
        /// Crear nueva pagina divisora
        /// </summary>
        /// <param name="mainpart"></param>
        /// <param name="title"></param>
        /// <param name="mar"></param>
        /// <param name="headerId"></param>
        /// <param name="footerId"></param>
        /// <param name="paSize"></param>
        public static void NewSDivider(ref MainDocumentPart mainpart, string title, (double, double, double, double) mar, string headerId = "", string footerId = "", PageSizeTypes paSize = PageSizeTypes.A4)
        {
            var res = CreateParagraph.NewP(title + StylesController.SetBold() + StylesController.SetFontSize(24), ParagraphTypes.Custom);
            mainpart.Document.Body.AppendChild(res);

            var p = NewS();
            var secProps1 = p.Descendants<SectionProperties>().FirstOrDefault();
            mainpart.Document.Body.AppendChild(p);


            var blanckHeader = mainpart.AddNewPart<HeaderPart>();
            var blanckHeaderPartId = mainpart.GetIdOfPart(blanckHeader);
            new Header().Save(blanckHeader);


            secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerId == "" ? blanckHeaderPartId : headerId });
            secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = footerId });
            secProps1.AppendChild(new VerticalTextAlignmentOnPage() { Val = VerticalJustificationValues.Center });

            SizeController.SetPageSize(secProps1, paSize, PageOrientationValues.Portrait);
            SizeController.SetMarginSize(secProps1, mar, PageOrientationValues.Portrait);
        }
    }
}
