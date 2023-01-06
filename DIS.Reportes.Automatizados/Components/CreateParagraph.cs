using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DIS.Reportes.Automatizados.Models.Enums;

namespace DIS.Reportes.Automatizados.Components
{
    public class CreateParagraph
    {
        /// <summary>
        /// Crear nuevo parrafo
        /// </summary>
        /// <param name="texto"></param>
        /// <param name="paragraphType"></param>
        /// <returns></returns>
        public static Paragraph NewP(string texto, ParagraphTypes paragraphType)
        {
            var paragraph = new Paragraph();
            var run = new Run();

            var runStyle = new StyleRunProperties();
            var paragraphStyle = new ParagraphProperties();
            SetParagraphStyles(paragraphType, ref runStyle, ref paragraphStyle, ref texto);

            var p = new Text(texto);
            p.Space = SpaceProcessingModeValues.Preserve;

            run.AppendChild(runStyle);
            run.AppendChild(p);
            paragraph.AppendChild(paragraphStyle);
            paragraph.AppendChild(run);

            return paragraph;
        }

        /// <summary>
        /// Crear nuevo parrafo que es salto de pagina
        /// </summary>
        /// <returns></returns>
        public static Paragraph NewPBreak()
        {
            return new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
        }


        #region Utils
        /// <summary>
        /// Establecer estilos acorde al tipo de parrafo dado
        /// </summary>
        /// <param name="paragraphType"></param>
        /// <param name="runStyle"></param>
        /// <param name="paragraphStyle"></param>
        /// <param name="text"></param>
        private static void SetParagraphStyles(ParagraphTypes paragraphType, ref StyleRunProperties runStyle, ref ParagraphProperties paragraphStyle, ref string text)
        {
            runStyle.AppendChild(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });

            switch (paragraphType)
            {
                case ParagraphTypes.Normal:
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "nn" };
                    break;

                case ParagraphTypes.Heading1:
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "Ttulo1" };
                    break;

                case ParagraphTypes.Heading2:
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "Ttulo2" };
                    break;

                case ParagraphTypes.Table:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Both });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" });

                    runStyle.AppendChild(new FontSize() { Val = "22" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Custom:
                    SetTextStyles(ref text, ref runStyle, ref paragraphStyle);
                    break;
            }
        }

        /// <summary>
        /// Establecer estilos personalizados de un parrafo en especifico
        /// </summary>
        /// <param name="texto"></param>
        /// <param name="rProps"></param>
        /// <param name="pProps"></param>
        private static void SetTextStyles(ref string texto, ref StyleRunProperties rProps, ref ParagraphProperties pProps)
        {
            var bold = texto.Contains("[N]");
            var italic = texto.Contains("[I]");
            var underline = texto.Contains("[U]");
            var fontSize = texto.Contains("[F:");
            var fontColor = texto.Contains("[FC:");
            var jLeft = texto.Contains("¬");

            if (bold)
            {
                rProps.AppendChild(new Bold());
                texto = texto.Replace("[N]", "");
            }

            if (italic)
            {
                rProps.AppendChild(new Italic());
                texto = texto.Replace("[I]", "");
            }

            if (underline)
            {
                rProps.AppendChild(new Underline() { Val = UnderlineValues.Single });
                texto = texto.Replace("[U]", "");
            }

            if (fontSize)
            {
                var fontSizeValue = texto.Substring(texto.IndexOf("[F:") + 3, 2);
                rProps.AppendChild(new FontSize() { Val = fontSizeValue });
                texto = texto.Replace("[F:" + fontSizeValue + "]", "");
            }

            if (fontColor)
            {
                var fontColorValue = texto.Substring(texto.IndexOf("[FC:") + 4, 7);
                rProps.AppendChild(new Color() { Val = fontColorValue });
                texto = texto.Replace("[FC:" + fontColorValue + "]", "");
            }
            else
            {
                rProps.AppendChild(new Color() { Val = "#000000" });
            }

            if (jLeft)
            {
                pProps.AppendChild(new Justification() { Val = JustificationValues.Left });
                texto = texto.Replace("¬", "");
            }
            else
            {
                pProps.AppendChild(new Justification() { Val = JustificationValues.Center });
            }
        }
        #endregion
    }
}
