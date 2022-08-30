using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml.ExtendedProperties;

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
    
    public class WordCommands
    {
        #region Crear y abrir documento
        public WordprocessingDocument CreateDocument(string route, bool autosave = true)
        {
            try
            {
                var document = WordprocessingDocument.Create(route, WordprocessingDocumentType.Document, autosave);
                return document;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public WordprocessingDocument OpenDocument(string route)
        {
            try
            {
                var document = WordprocessingDocument.Open(route, true);
                return document;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
        #endregion


        #region Crear imagenes
        public Paragraph CreateNewBase64Image(MainDocumentPart mainPart, string base64, double escale = 1)
        {
            try
            {
                ImagePart imagePart;

                if (base64.ToLower().Contains("[jpeg]"))
                {
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    base64 = base64.Replace("[jpeg]", "");
                }
                else if (base64.ToLower().Contains("[png]"))
                {
                    imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    base64 = base64.Replace("[png]", "");
                }
                else throw new Exception("Formato de imagen no soportado");

                
                byte[] imgBytes = Convert.FromBase64String(base64);
                BitmapImage img = new BitmapImage();
                
                img.BeginInit();
                img.StreamSource = new MemoryStream(imgBytes);
                img.EndInit();
                imagePart.FeedData(new MemoryStream(imgBytes));

                const int emusPerInch = 914400;
                var wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                var hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(wImgEmus * escale, hImgEmus * escale));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        public Paragraph CreateNewBase64Image(MainDocumentPart mainPart, string base64, long width = 0, long height = 0)
        {
            try
            {
                ImagePart imagePart;

                if (base64.ToLower().Contains("[jpeg]"))
                {
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    base64 = base64.Replace("[jpeg]", "");
                }
                else if (base64.ToLower().Contains("[png]"))
                {
                    imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    base64 = base64.Replace("[png]", "");
                }
                else throw new Exception("Formato de imagen no soportado");

                
                byte[] imgBytes = Convert.FromBase64String(base64);
                BitmapImage img = new BitmapImage();

                img.BeginInit();
                img.StreamSource = new MemoryStream(imgBytes);
                img.EndInit();
                imagePart.FeedData(new MemoryStream(imgBytes));

                const int emusPerInch = 914400;
                const int emusPerCm = 360000;

                var wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                var hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);
                var wDifined = (long)(width * emusPerCm);
                var hDefined = (long)(height * emusPerCm);

                if (width == 0)
                {
                    var ratio = (wImgEmus * 1.0m) / hImgEmus;
                    wDifined = (long)(hDefined * ratio);
                }
                else if (height == 0)
                {
                    var ratio = (hImgEmus * 1.0m) / wImgEmus;
                    hDefined = (long)(wDifined * ratio);
                }

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(wDifined, hDefined));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        public Paragraph CreateNewImage(MainDocumentPart mainPart, string fileName, double escale = 1)
        {
            try
            {
                ImagePart imagePart;

                if (fileName.ToLower().Contains(".jpeg") || fileName.ToLower().Contains(".jpg"))
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                else if (fileName.ToLower().Contains(".png"))
                    imagePart = mainPart.AddImagePart(ImagePartType.Png);

                else throw new Exception("Formato de imagen no soportado");


                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                    stream.Close();
                }

                var img = new BitmapImage();
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    img.BeginInit();
                    img.StreamSource = fs;
                    img.EndInit();
                }

                const int emusPerInch = 914400;
                var wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                var hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(wImgEmus * escale, hImgEmus * escale));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }
        
        public Paragraph CreateNewImage(MainDocumentPart mainPart, string fileName, long width = 0, long height = 0)
        {
            try
            {
                ImagePart imagePart;

                if (fileName.ToLower().Contains(".jpeg") || fileName.ToLower().Contains(".jpg"))
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                else if (fileName.ToLower().Contains(".png"))
                    imagePart = mainPart.AddImagePart(ImagePartType.Png);

                else throw new Exception("Formato de imagen no soportado");


                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                    stream.Close();
                }

                var img = new BitmapImage();
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    img.BeginInit();
                    img.StreamSource = fs;
                    img.EndInit();
                }

                const int emusPerInch = 914400;
                const int emusPerCm = 360000;

                var wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                var hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);
                var wDifined = (long)(width * emusPerCm);
                var hDefined = (long)(height * emusPerCm);
                
                if (width == 0)
                {
                    var ratio = (wImgEmus * 1.0m) / hImgEmus;
                    wDifined = (long)(hDefined * ratio);
                }
                else if (height == 0)
                {
                    var ratio = (hImgEmus * 1.0m) / wImgEmus;
                    hDefined = (long)(wDifined * ratio);
                }

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(wDifined, hDefined));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        private static Paragraph CreateNewImageElement(string relationId, Size size)
        {
            Int64Value width = (Int64Value)size.Width;
            Int64Value height = (Int64Value)size.Height;

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width, Cy = height },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = relationId + " image name"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = relationId
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()
                                    {
                                        PictureLocks = new A.PictureLocks() { NoChangeAspect = true }
                                    }
                                ),

                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                                        )
                                    )
                                    {
                                        Embed = relationId,
                                        CompressionState =
                                        A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle()
                                    )
                                ),

                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = width, Cy = height }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                }
            );

            return new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), new Run(element));
        }
        #endregion


        #region Crear secciones
        public void CreateNewSectionDivider(ref MainDocumentPart mainpart, string title, (double, double, double, double) mar, string headerId = "", string footerId = "", PageSizeTypes paSize = PageSizeTypes.A4)
        {
            var res = CreateNewParagraph(title, ParagraphTypes.Heading1);
            mainpart.Document.Body.AppendChild(res);

            var p = CreateNewSection();
            var secProps1 = p.Descendants<SectionProperties>().FirstOrDefault();
            mainpart.Document.Body.AppendChild(p);


            var blanckHeader = mainpart.AddNewPart<HeaderPart>();
            var blanckHeaderPartId = mainpart.GetIdOfPart(blanckHeader);
            new Header().Save(blanckHeader);


            secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerId == "" ? blanckHeaderPartId : headerId });
            secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = footerId });
            secProps1.AppendChild(new VerticalTextAlignmentOnPage() { Val = VerticalJustificationValues.Center });
            
            WordUtils.SetPageSize(secProps1, paSize, PageOrientationValues.Portrait);
            WordUtils.SetMarginSize(secProps1, mar, PageOrientationValues.Portrait);
        }

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
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.NextPage });

            return SectionBreakProperties;
        }
        #endregion


        #region Crear footer y header
        public Header CreateNewHeaderForSection(string pretitle, string title)
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

        public Footer CreateNewFooterForSection(string footerText)
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


        #region Crear tablas
        public Table CreateNewTable(List<string[]> datosTabla, bool haveBorder = true)
        {
            Table table = new Table(new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new BottomBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new LeftBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new RightBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new InsideHorizontalBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new InsideVerticalBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    }
                ),
                new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                )
            ));

            int rowCount = datosTabla.Count;
            int columnCount = GetColsNumber(datosTabla);

            for (int row = 0; row < rowCount; row++)
            {
                TableRow tableRow = new TableRow();

                for (int col = 0; col < columnCount; col++)
                {
                    var texto = datosTabla[row].Length > col ? datosTabla[row][col] : "";

                    var pProps = new ParagraphProperties(
                        new SpacingBetweenLines() { Before = "40", After = "40" },
                        new Indentation() { Left = "40" },
                        new Languages() { Val = "es-ES" }
                    );
                    var rProps = new RunProperties(
                        new FontSize() { Val = "20" },
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }
                    );
                    var cProps = new TableCellProperties(
                        new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                    );


                    SetCellTextStyles(ref texto, ref rProps, ref pProps);

                    #region Merge celdas
                    // "~" caracter que indica unir las dos celdas horizontalmente
                    // "|" caracter que indica unir las dos celdas verticalmente

                    // Validar si la celda es una celda de unir horizontalmente
                    var cellMerge = texto.Contains("~");
                    if (cellMerge)
                    {
                        texto = texto.Replace("~", "");
                        cProps.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cProps.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                    }

                    // Validar si la celda es una celda de unir verticalmente
                    var rowMerge = texto.Contains("|");
                    if (rowMerge)
                    {
                        texto = texto.Replace("|", "");
                        cProps.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cProps.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                    }
                    #endregion


                    TableCell tableCell = new TableCell(
                        cProps,
                        new Paragraph(
                            pProps,
                            new Run(
                                rProps,
                                new Text(texto)
                            )
                        )
                    );

                    tableRow.Append(tableCell);
                }

                table.Append(tableRow);
            }

            return table;
        }

        public Table CreateNewImageTable(List<string[]> datosTabla, (double width, double height) utilSpace, MainDocumentPart mainPart, bool haveBorder = true)
        {
            Table table = new Table(new TableProperties(
                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new BottomBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new LeftBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new RightBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new InsideHorizontalBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    },
                    new InsideVerticalBorder()
                    {
                        Val = haveBorder ? BorderValues.Single : BorderValues.None,
                        Size = haveBorder ? (UInt32Value)10 : 0,
                    }
                ),
                new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = "0", Type = TableWidthUnitValues.Dxa }
                )
            ));

            int rowCount = datosTabla.Count;
            int columnCount = GetColsNumber(datosTabla);

            for (int row = 0; row < rowCount; row++)
            {
                TableRow tableRow = new TableRow();

                for (int col = 0; col < columnCount; col++)
                {
                    var texto = datosTabla[row].Length > col ? datosTabla[row][col] : "";

                    var pProps = new ParagraphProperties(
                        new SpacingBetweenLines() { Before = "40", After = "40" },
                        new Indentation() { Left = "40" },
                        new Languages() { Val = "es-ES" }
                    );
                    var rProps = new RunProperties(
                        new FontSize() { Val = "20" },
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }
                    );
                    var cProps = new TableCellProperties(
                        new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                    );


                    SetCellTextStyles(ref texto, ref rProps, ref pProps);


                    #region Merge celdas
                    // "~" caracter que indica unir las dos celdas horizontalmente
                    // "|" caracter que indica unir las dos celdas verticalmente

                    // Validar si la celda es una celda de unir horizontalmente
                    var cellMerge = texto.Contains("~");
                    if (cellMerge)
                    {
                        texto = texto.Replace("~", "");
                        cProps.Append(new HorizontalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cProps.Append(new HorizontalMerge() { Val = MergedCellValues.Restart });
                    }

                    // Validar si la celda es una celda de unir verticalmente
                    var rowMerge = texto.Contains("|");
                    if (rowMerge)
                    {
                        texto = texto.Replace("|", "");
                        cProps.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                    }
                    else
                    {
                        cProps.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
                    }
                    #endregion



                    if (File.Exists(texto))
                    {
                        var img = CreateNewImage(mainPart, texto, width: (long)(utilSpace.width / 2));
                        TableCell tableCell = new TableCell(
                            cProps,
                            img
                        );

                        tableRow.Append(tableCell);
                    }
                    else
                    {
                        TableCell tableCell = new TableCell(
                            cProps,
                            new Paragraph(
                                pProps,
                                new Run(
                                    rProps,
                                    new Text(texto)
                                )
                            )
                        );

                        tableRow.Append(tableCell);
                    }
                }

                table.Append(tableRow);
            }

            return table;
        }

        private static void SetCellTextStyles(ref string texto, ref RunProperties rProps, ref ParagraphProperties pProps)
        {
            var bold = texto.Contains("[N]");
            var italic = texto.Contains("[I]");
            var underline = texto.Contains("[U]");
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

        private static int GetColsNumber(List<string[]> DatosTabla)
        {
            int Col = 0;

            for (int i = 0; i < DatosTabla.Count; i++)
            {
                if (Col < DatosTabla[i].Length)
                {
                    Col = DatosTabla[i].Length;
                }
            }
            return Col;
        }
        #endregion


        #region Crear parrafos
        public Paragraph CreateNewParagraph(string texto, ParagraphTypes paragraphType)
        {
            var paragraph = new Paragraph();
            var run = new Run();

            var runStyle = new StyleRunProperties();
            var paragraphStyle = new ParagraphProperties();
            SetParagraphStyles(paragraphType, ref runStyle, ref paragraphStyle);

            var p = new Text(texto);
            p.Space = SpaceProcessingModeValues.Preserve;

            run.AppendChild(runStyle);
            run.AppendChild(p);
            paragraph.AppendChild(paragraphStyle);
            paragraph.AppendChild(run);

            return paragraph;
        }

        public Paragraph CreateNewPargraphPageBreak()
        {
            return new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
        }

        private static void SetParagraphStyles(ParagraphTypes paragraphType, ref StyleRunProperties runStyle, ref ParagraphProperties paragraphStyle)
        {
            runStyle.AppendChild(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });

            switch (paragraphType)
            {
                case ParagraphTypes.Normal:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Both });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "160", After = "160" });

                    runStyle.AppendChild(new FontSize() { Val = "24" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Heading1:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Center });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "160", After = "160" });

                    runStyle.AppendChild(new Bold());
                    runStyle.AppendChild(new FontSize() { Val = "24" });
                    runStyle.AppendChild(new Color() { Val = "#000000" });
                    break;

                case ParagraphTypes.Heading2:
                    paragraphStyle.AppendChild(new Justification() { Val = JustificationValues.Left });
                    paragraphStyle.AppendChild(new SpacingBetweenLines() { LineRule = LineSpacingRuleValues.Auto, Before = "160", After = "160" });

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
        }
        #endregion
    }
}
