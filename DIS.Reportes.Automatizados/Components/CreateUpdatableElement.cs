using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace DIS.Reportes.Automatizados.Components
{
    public class CreateUpdatableElement
    {
        #region Create Header and Footer
        public static Header NewHeader(string title, string pretitle)
        {
            Header header = new Header();

            #region NameSpaces
            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            #endregion

            Table headerTable = new Table(new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },

                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                ),
                new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                )
            ));


            TableRow headerRow1 = new TableRow();
            TableRow headerRow2 = new TableRow();

            TableCell headerCell11 = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Left },
                        new SpacingBetweenLines() { Before = "0", After = "22" },
                        new Languages() { Val = "es-ES" }
                    ),
                    new Run(
                        new RunProperties(new FontSize() { Val = "20" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                        new Text(pretitle)
                    )
                )
            );
            TableCell headerCell21 = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Left },
                        new SpacingBetweenLines() { Before = "0", After = "22" },
                        new Languages() { Val = "es-ES" }
                    ),
                    new Run(
                        new RunProperties(new FontSize() { Val = "20" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                        new Text(title)
                    )
                )
            );
            TableCell headerCell22 = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Right },
                        new SpacingBetweenLines() { Before = "0", After = "22" },
                        new Languages() { Val = "es-ES" }
                    ),
                    new Run(
                        new RunProperties(new FontSize() { Val = "20" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                        new Text("Página ") { Space = SpaceProcessingModeValues.Preserve },
                        new SimpleField() { Instruction = "PAGE" },
                        new Text(" de ") { Space = SpaceProcessingModeValues.Preserve },
                        new SimpleField() { Instruction = "NUMPAGES" }
                    )
                )
            );

            headerRow1.Append(headerCell11);
            headerRow2.Append(headerCell21, headerCell22);

            headerTable.Append(headerRow1, headerRow2);
            header.Append(headerTable);

            return header;
        }

        public static Footer NewFooter(string footerText)
        {
            Footer footer = new Footer();

            #region NameSpaces
            footer.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            #endregion

            Table footerTable = new Table(new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 10 },

                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                ),
                new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                )
            ));

            TableRow footerRow1 = new TableRow();
            TableCell footerCell11 = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Right },
                        new SpacingBetweenLines() { Before = "0", After = "22" },
                        new Languages() { Val = "es-ES" }
                    ),
                    new Run(
                        new RunProperties(new FontSize() { Val = "20" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                        new Text(footerText)
                    )
                )
            );

            footerRow1.Append(footerCell11);
            footerTable.Append(footerRow1);
            footer.Append(footerTable);

            return footer;
        }
        #endregion

        #region Create TOCs
        /// <summary>
        /// Crear una nueva tabla de contenidos que liste los titulos principales
        /// </summary>
        /// <param name="doc"></param>
        public static void NewTOC(Document doc)
        {
            var sdtBlock = new SdtBlock();
            var title = "TABLA DE CONTENIDO";

            sdtBlock.InnerXml = $@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
              <w:sdtPr>
                <w:docPartObj>
                  <w:docPartGallery w:val='Table of Contents'/>
                  <w:docPartUnique/>
                </w:docPartObj>
              </w:sdtPr>
              <w:sdtEndPr>
                <w:rPr>
                 <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                 <w:color w:val='auto'/>
                 <w:sz w:val='22'/>
                 <w:szCs w:val='22'/>
                 <w:lang w:eastAsia='en-US'/>
                </w:rPr>
              </w:sdtEndPr>
              <w:sdtContent>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='TOCHeading'/>
                  </w:pPr>
                  <w:r>
                    <w:t>{title}</w:t>
                  </w:r>
                </w:p>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='TOC1'/>
                    <w:tabs>
                      <w:tab w:val='right' w:leader='dot'/>
                    </w:tabs>
                    <w:rPr>
                      <w:noProof/>
                    </w:rPr>
                  </w:pPr>
                  <w:r>
                    <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                  </w:r>
                  <w:r>
                    <w:instrText xml:space='preserve'>TOC \o ""1-3"" \h \z \u</w:instrText>
                  </w:r>
                  <w:r>
                    <w:fldChar w:fldCharType='separate'/>
                  </w:r>
                </w:p>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType='end'/>
                  </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>";
            doc.MainDocumentPart.Document.Body.AppendChild(sdtBlock);

            var settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings { BordersDoNotSurroundFooter = new BordersDoNotSurroundFooter() { Val = true } };
            settingsPart.Settings.Append(new UpdateFieldsOnOpen() { Val = true });
        }
        #endregion
    }
}
