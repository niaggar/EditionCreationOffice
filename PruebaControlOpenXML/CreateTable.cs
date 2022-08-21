using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PruebaControlOpenXML
{
    public enum ParagraphTypes
    {
        Normal,
        Heading1,
        Heading2,
        Table,
    }

    public enum PageSizeTypes
    {
        A4,
        A5,
        A6,
        Custom,
    }
    
    public class CreateTable
    {
        public WordprocessingDocument CrearNuevoDocumento(string route)
        {
            var document = WordprocessingDocument.Create(@"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\TestOpenXML.docx", WordprocessingDocumentType.Document);

            return document;
        }

        
        #region Crear secciones
        public Paragraph CreateNewSection()
        {
            var paragraphSectionBreak = new Paragraph();
            var paragraphSectionBreakProperties = new ParagraphProperties();
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.NextPage });

            paragraphSectionBreakProperties.Append(SectionBreakProperties);
            paragraphSectionBreak.Append(paragraphSectionBreakProperties);

            return paragraphSectionBreak;
        }

        public SectionProperties CreateFinalSection()
        {
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.Continuous });

            return SectionBreakProperties;
        }
        #endregion


        #region Crear footer y header
        public Header CreateHeaderForSection(string pretitle, string title)
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
                        new RunProperties(new FontSize() { Val = "22" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
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
                        new RunProperties(new FontSize() { Val = "22" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
                        new Text(title)
                    )
                )
            );
            TableCell headerCell22 = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Left },
                        new SpacingBetweenLines() { Before = "0", After = "22" },
                        new Languages() { Val = "es-ES" }
                    ),
                    new Run(
                        new RunProperties(new FontSize() { Val = "22" }, new Bold(), new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
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

        public Footer CreateFooterForSection(string footerText)
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
                        new RunProperties(new FontSize() { Val = "22" }, new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }),
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





        public Paragraph CrearNuevoParrafo(string texto, ParagraphTypes paragraphType)
        {
            var paragraph = new Paragraph();
            var run = new Run();
            
            var runStyle = new StyleRunProperties();
            var paragraphStyle = new ParagraphProperties();
            AgregarEstilosDeParrafo(paragraphType, ref runStyle, ref paragraphStyle);

            var p = new Text(texto);
            p.Space = SpaceProcessingModeValues.Preserve;

            run.AppendChild(runStyle);
            run.AppendChild(p);
            paragraph.AppendChild(paragraphStyle);
            paragraph.AppendChild(run);

            return paragraph;
        }

        public Table CrearNuevaTablaWord()
        {
            var table = new Table();
            var tableProperties = new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 }
                ),
                new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                )
            );

            table.AppendChild(tableProperties);

            

            TableRow tr = new TableRow();
            TableCell cell1 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));
            TableCell cell2 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));
            TableCell cell3 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));

            TableCellProperties cellprops1 = new TableCellProperties(new HorizontalMerge() { Val = MergedCellValues.Restart });
            TableCellProperties cellprops2 = new TableCellProperties(new HorizontalMerge() { Val = MergedCellValues.Continue });

            cell1.Append(cellprops1);
            cell2.Append(cellprops2);

            tr.Append(cell1, cell2, cell3);

                
            TableRow tr1 = new TableRow();
            TableCell cell4 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));
            TableCell cell5 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));
            TableCell cell6 = new TableCell(CrearNuevoParrafo("Hola", ParagraphTypes.Table));
            tr1.Append(cell4, cell5, cell6);


            table.AppendChild(tr);
            table.AppendChild(tr1);


            return table;
        }


        public void AgregarEstilosDeParrafo(ParagraphTypes paragraphType, ref StyleRunProperties runStyle, ref ParagraphProperties paragraphStyle)
        {
            runStyle.AppendChild(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });

            switch (paragraphType)
            {
                case ParagraphTypes.Normal:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Both });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "120", After = "120" });

                    runStyle.AppendChild(new FontSize() { Val = "24" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Heading1:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Center });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "120", After = "120" });

                    runStyle.AppendChild(new Bold());
                    runStyle.AppendChild(new FontSize() { Val = "24" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Heading2:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Left });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "120", After = "120" });

                    runStyle.AppendChild(new Bold());
                    runStyle.AppendChild(new Italic());
                    runStyle.AppendChild(new FontSize() { Val = "24" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Table:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Both });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" });

                    runStyle.AppendChild(new FontSize() { Val = "22" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;
            }
            
            #region propiedades
            //Bold bold1 = new Bold();
            //Italic italic1 = new Italic();
            //Underline underline = new Underline();

            //Color color1 = new Color() { Val = "#FF0000" };
            //RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            //FontSize fontSize1 = new FontSize() { Val = "24" };
            #endregion
        }
    }
}
