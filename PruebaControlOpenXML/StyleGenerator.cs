using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace PruebaControlOpenXML
{
    public class StyleGenerator
    {
        public StyleGenerator() { }

        public static void CreateAndAddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart)
        {
            Styles styles = styleDefinitionsPart.Styles;
            if (styles == null)
            {
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();
            }


            #region Heading 1
            // Create a new paragraph style element and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "tt1",
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "tt1Car" };
            PrimaryStyle primaryStyle = new PrimaryStyle();
            ParagraphProperties pprops = new ParagraphProperties();
            NumberingProperties nprop = new NumberingProperties();
            SpacingBetweenLines space1 = new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "240", After = "240" };
            Justification just1 = new Justification() { Val = JustificationValues.Both };
            OutlineLevel outl1= new OutlineLevel() { Val = 0 };
            StyleName styleName1 = new StyleName() { Val = "tt1" };

            nprop.Append(new NumberingId() { Val = 15 });
            nprop.Append(new NumberingLevelReference() { Val = 0 });
            pprops.Append(nprop);
            pprops.Append(space1);
            pprops.Append(just1);
            pprops.Append(outl1);

            style.Append(basedon1);
            style.Append(linkedStyle1);
            style.Append(nextParagraphStyle1);
            style.Append(primaryStyle);
            style.Append(pprops);
            style.Append(styleName1);


            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { Val = "#000000" };
            RunFonts font1 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic1 = new Italic();
            FontSize fontSize1 = new FontSize() { Val = "24" };


            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);
            #endregion

            #region Heading 2
            // Create a new paragraph style element and specify some of the attributes.
            Style style2 = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "tt2",
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon2 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "tt2Car" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            ParagraphProperties pprops2 = new ParagraphProperties();
            NumberingProperties nprop2 = new NumberingProperties();
            StyleName styleName2 = new StyleName() { Val = "tt2" };
            SpacingBetweenLines space2 = new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "240", After = "120" };
            Justification just2 = new Justification() { Val = JustificationValues.Both };
            OutlineLevel outl2 = new OutlineLevel() { Val = 1 };

            nprop2.Append(new NumberingId() { Val = 15 });
            nprop2.Append(new NumberingLevelReference() { Val = 1 });
            pprops2.Append(nprop2);
            pprops2.Append(space2);
            pprops2.Append(just2);
            pprops2.Append(outl2);

            style2.Append(basedon2);
            style2.Append(linkedStyle2);
            style2.Append(nextParagraphStyle2);
            style2.Append(primaryStyle2);
            style2.Append(pprops2);
            style2.Append(styleName2);


            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            Bold bold2 = new Bold();
            Color color2 = new Color() { Val = "#000000" };
            RunFonts font2 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic2 = new Italic();
            FontSize fontSize2 = new FontSize() { Val = "24" };


            styleRunProperties2.Append(bold2);
            styleRunProperties2.Append(color2);
            styleRunProperties2.Append(font2);
            styleRunProperties2.Append(fontSize2);
            styleRunProperties2.Append(italic2);

            // Add the run properties to the style.
            style2.Append(styleRunProperties2);
            #endregion

            #region Normal
            // Create a new paragraph style element and specify some of the attributes.
            Style style3 = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "nn",
                CustomStyle = true,
                Default = true
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon3 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "Normal" };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            ParagraphProperties pprops3 = new ParagraphProperties();
            SpacingBetweenLines space3 = new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "160", After = "160" };
            Justification just3 = new Justification() { Val = JustificationValues.Both };
            StyleName styleName3 = new StyleName() { Val = "nn" };

            pprops3.Append(space3);
            pprops3.Append(just3);

            style3.Append(basedon3);
            style3.Append(nextParagraphStyle3);
            style3.Append(primaryStyle3);
            style3.Append(pprops3);
            style3.Append(styleName3);


            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            Color color3 = new Color() { Val = "#000000" };
            RunFonts font3 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize3 = new FontSize() { Val = "24" };


            styleRunProperties3.Append(color3);
            styleRunProperties3.Append(font3);
            styleRunProperties3.Append(fontSize3);

            // Add the run properties to the style.
            style3.Append(styleRunProperties3);
            #endregion


            styles.Append(style);
            styles.Append(style2);
            styles.Append(style3);
        }

        public static StyleDefinitionsPart AddStylesPartToPackage(Document doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            Styles root = new Styles();
            root.Save(part);

            return part;
        }
    }
}
