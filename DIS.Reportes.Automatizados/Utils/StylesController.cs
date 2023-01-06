using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIS.Reportes.Automatizados.Utils
{
    public class StylesController
    {
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
                StyleId = "Ttulo1",
                CustomStyle = false,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "nn" };
            Aliases aliases1 = new Aliases() { Val = "UCI Header 1,CONT,CVRD,Título 0,Título 1_HTA,Edgar 1,oscar1,GT Título 1,HT-IF Título 1" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Ttulo1Car" };
            PrimaryStyle primaryStyle = new PrimaryStyle();
            ParagraphProperties pprops = new ParagraphProperties();
            NumberingProperties nprop = new NumberingProperties();
            SpacingBetweenLines space1 = new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "240", After = "240" };
            Justification just1 = new Justification() { Val = JustificationValues.Both };
            OutlineLevel outl1 = new OutlineLevel() { Val = 0 };
            StyleName styleName1 = new StyleName() { Val = "heading 1" };

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
            style.Append(aliases1);


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
                StyleId = "Ttulo2",
                CustomStyle = false,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon2 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "nn" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Ttulo2Car" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Aliases aliases2 = new Aliases() { Val = "GT Título 2,HT-IF Título 2" };
            ParagraphProperties pprops2 = new ParagraphProperties();
            NumberingProperties nprop2 = new NumberingProperties();
            StyleName styleName2 = new StyleName() { Val = "heading 2" };
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
            style2.Append(aliases2);


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
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "nn" };
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

            #region TOC normal
            Style style4 = new Style();
            style4.InnerXml = @"
            <w:style w:type='paragraph' w:styleId='TOCHeading' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                <w:name w:val='TOC Heading'/>
                <w:basedOn w:val='Heading1'/>
                <w:next w:val='Normal'/>
                <w:uiPriority w:val='39'/>
                <w:semiHidden/>
                <w:unhideWhenUsed/>
                <w:qFormat/>
                <w:pPr>
                    <w:jc w:val=""center"" />
                    <w:outlineLvl w:val='9'/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial"" />
                    <w:b />
                    <w:color w:val=""000000"" />
                    <w:sz w:val=""24"" />
                </w:rPr>
            </w:style>";

            Style style5 = new Style();
            style5.InnerXml = @"
            <w:style w:type='paragraph' w:styleId='TOC1' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                <w:name w:val='toc 1'/>
                <w:basedOn w:val='Normal'/>
                <w:next w:val='Normal'/>
                <w:autoRedefine/>
                <w:uiPriority w:val='39'/>
                <w:unhideWhenUsed/>
                <w:pPr>
                    <w:ind w:left=""0"" w:right=""0"" w:hanging=""0"" />  
                    <w:spacing w:after='100'/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial"" />
                    <w:b />
                    <w:color w:val=""#000000"" />
                    <w:sz w:val=""20"" />
                </w:rPr>
            </w:style>";

            Style style6 = new Style();
            style6.InnerXml = @"
            <w:style w:type='paragraph' w:styleId='TOC2' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                <w:name w:val='toc 2'/>
                <w:basedOn w:val='Normal'/>
                <w:next w:val='Normal'/>
                <w:autoRedefine/>
                <w:uiPriority w:val='39'/>
                <w:unhideWhenUsed/>
                <w:pPr>
                    <w:ind w:left=""244"" w:right=""0"" w:hanging=""0"" />  
                    <w:spacing w:after='100'/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial"" />
                    <w:b />
                    <w:color w:val=""#000000"" />
                    <w:sz w:val=""20"" />
                </w:rPr>
            </w:style>";


            Style style7 = new Style();
            style7.InnerXml = @"
            <w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                <w:name w:val='Hyperlink'/>
                <w:basedOn w:val='DefaultParagraphFont'/>
                <w:uiPriority w:val='99'/>
                <w:unhideWhenUsed/>
                <w:rPr>
                    <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial"" />
                    <w:b />
                    <w:color w:val=""#000000"" />
                    <w:sz w:val=""20"" />
                </w:rPr>
            </w:style>";
            #endregion

            #region Caption
            Style style8 = new Style();
            style8.InnerXml = @"
            <w:style w:type=""paragraph"" w:styleId=""Caption"">
                <w:name w:val=""Caption"" />
                <w:aliases
                w:val=""EpiTabla,Epígrafe3,Epígrafe1-Figura,Epígrafe Car1,IEB Tabla eprigrafe,Car,Epigrafe,Epígrafe Car21,Descripción1,Foto"" />
                <w:basedOn w:val=""Normal"" />
                <w:next w:val=""Normal"" />
                <w:link w:val=""DescripcinCar1"" />
                <w:qFormat />
                <w:rsid w:val=""009A120D"" />
                <w:pPr>
                    <w:keepNext />
                    <w:widowControl w:val=""0"" />
                    <w:spacing w:before=""60"" />
                    <w:jc w:val=""center"" />
                </w:pPr>
                <w:rPr>
                    <w:b />
                    <w:sz w:val=""20"" />
                    <w:lang w:val=""es-CL"" />
                </w:rPr>
            </w:style>";
            #endregion


            styles.Append(style);
            styles.Append(style2);
            styles.Append(style3);
            styles.Append(style4);
            styles.Append(style5);
            styles.Append(style6);
            styles.Append(style7);
            styles.Append(style8);
        }

        public static void AddNumberingPart(Document doc)
        {
            var numbering = doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

            Numbering globalNumbering = new Numbering();

            #region Namespaces
            globalNumbering.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            globalNumbering.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            globalNumbering.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            globalNumbering.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            globalNumbering.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            globalNumbering.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            globalNumbering.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            globalNumbering.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            globalNumbering.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            globalNumbering.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            globalNumbering.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            globalNumbering.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            globalNumbering.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            globalNumbering.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            globalNumbering.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            globalNumbering.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            globalNumbering.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            globalNumbering.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            globalNumbering.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            globalNumbering.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            #endregion

            var abs = new AbstractNum() { AbstractNumberId = 0 };
            var nsid3 = new Nsid() { Val = "70913756" };
            var multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            var templateCode3 = new TemplateCode() { Val = "624EA66A" };
            var abstractNumDefinitionName1 = new AbstractNumDefinitionName() { Val = "TitlesNumberingDEFAULT" };

            #region Level 1 of the list
            var level1 = new Level() { LevelIndex = 0 };

            var pPr1 = new PreviousParagraphProperties();
            pPr1.Append(new Tabs(new TabStop() { Val = TabStopValues.Number, Position = 0 }));
            pPr1.Append(new Indentation() { Left = "432", Hanging = "432" });

            var rPr1 = new NumberingSymbolRunProperties();
            rPr1.Append(new RunFonts() { Hint = FontTypeHintValues.Default });

            level1.Append(new StartNumberingValue() { Val = 1 });
            level1.Append(new NumberingFormat() { Val = NumberFormatValues.Decimal });
            level1.Append(new LevelText() { Val = "%1" });
            level1.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
            level1.Append(new ParagraphStyleIdInLevel() { Val = "Ttulo1" });
            level1.Append(pPr1);
            level1.Append(rPr1);
            #endregion

            #region Level 2 of the list
            var level2 = new Level() { LevelIndex = 1 };

            var pPr2 = new PreviousParagraphProperties();
            pPr2.Append(new Tabs(new TabStop() { Val = TabStopValues.Number, Position = 0 }));
            pPr2.Append(new Indentation() { Left = "576", Hanging = "576" });

            var rPr2 = new NumberingSymbolRunProperties();
            rPr2.Append(new RunFonts() { Hint = FontTypeHintValues.Default });

            level2.Append(new NumberingFormat() { Val = NumberFormatValues.Decimal });
            level2.Append(new ParagraphStyleIdInLevel() { Val = "Ttulo2" });
            level2.Append(new StartNumberingValue() { Val = 1 });
            level2.Append(new LevelText() { Val = "%1.%2" });
            level2.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
            level2.Append(pPr2);
            level2.Append(rPr2);
            #endregion

            // Add the props and levels to the abstract of the list
            abs.Append(nsid3);
            abs.Append(multiLevelType3);
            abs.Append(templateCode3);
            abs.Append(abstractNumDefinitionName1);
            abs.Append(level1);
            abs.Append(level2);

            #region Rest of the levels
            for (int i = 2; i < 9; i++)
            {
                var level = new Level() { LevelIndex = i };

                var pPr = new PreviousParagraphProperties();
                pPr.Append(new Tabs(new TabStop() { Val = TabStopValues.Number, Position = 0 }));
                pPr.Append(new Indentation() { Left = $"{577 + (144 * (i - 1))}", Hanging = $"{577 + (144 * (i - 1))}" });

                var rPr = new NumberingSymbolRunProperties();
                rPr.Append(new RunFonts() { Hint = FontTypeHintValues.Default });

                level.Append(new NumberingFormat() { Val = NumberFormatValues.Decimal });
                level.Append(new ParagraphStyleIdInLevel() { Val = $"tt{i + 1}" });
                level.Append(new StartNumberingValue() { Val = 1 });
                level.Append(new LevelText() { Val = $"%{i}" });
                level.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
                level.Append(pPr);
                level.Append(rPr);

                abs.Append(level);
            }
            #endregion

            // Create the new list with the abstract
            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = 1 };
            numberingInstance.Append(new AbstractNumId() { Val = 0 });
            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 15 };
            numberingInstance.Append(new AbstractNumId() { Val = 0 });
            globalNumbering.Append(abs);
            globalNumbering.Append(numberingInstance);
            globalNumbering.Append(numberingInstance2);

            // Save the created list
            numbering.Numbering = globalNumbering;
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
