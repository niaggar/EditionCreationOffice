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
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

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

        public Paragraph CreateNewBase64Image(MainDocumentPart mainPart, string base64, double width = 0, double height = 0)
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
                var extension = System.IO.Path.GetExtension(fileName);
                if (extension == ".jpeg" || extension == ".jpg")
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                else if (extension == ".png")
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












        // Metodos corregidos para generar imagenes con el tamaño especifico dado como parametro

        /// <summary>
        /// Crear imagenes con width y height definidos
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public Paragraph CreateNewImage(MainDocumentPart mainPart, string fileName, double width = 0, double height = 0)
        {
            try
            {
                ImagePart imagePart;
                var extension = System.IO.Path.GetExtension(fileName);
                if (extension == ".jpeg" || extension == ".jpg")
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                else if (extension == ".png")
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

                long emuWidth = (long)(height * 914400);
                long emuHeight = (long)(width * 914400);

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuWidth, emuHeight));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        /// <summary>
        /// Crear imagenes con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public Paragraph CreateNewImageWidth(MainDocumentPart mainPart, string fileName, double width, double maxheight, bool scale = true)
        {
            try
            {
                ImagePart imagePart;
                var extension = System.IO.Path.GetExtension(fileName);
                if (extension == ".jpeg" || extension == ".jpg")
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                else if (extension == ".png")
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



                long emuWidth = (long)(width * 914400);
                long maxEmuHeight = (long)(maxheight * 914400);

                long wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                long hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);

                var ratHeight = hImgEmus / wImgEmus;
                long emuHeight = (long)(emuWidth * ratHeight);


                if (maxheight != 0 && maxEmuHeight < emuHeight && scale == true)
                {
                    emuHeight = maxEmuHeight;
                    var ratWidht = wImgEmus / hImgEmus;
                    emuWidth = (long)(maxEmuHeight * ratWidht);
                }
                else if (maxheight != 0 && maxEmuHeight < emuHeight && scale == false)
                {
                    emuHeight = maxEmuHeight;
                }

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuWidth, emuHeight));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        /// <summary>
        /// Crear imagenes con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public Paragraph CreateNewImageHeight(MainDocumentPart mainPart, string fileName, double height, double maxwidth, bool scale = true)
        {
            try
            {
                ImagePart imagePart;
                var extension = System.IO.Path.GetExtension(fileName);
                if (extension == ".jpeg" || extension == ".jpg")
                    imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                else if (extension == ".png")
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



                long emuHeight = (long)(height * 914400);
                long maxEmuWidth = (long)(maxwidth * 914400);

                long wImgEmus = (long)(img.PixelWidth / img.DpiX * emusPerInch);
                long hImgEmus = (long)(img.PixelHeight / img.DpiY * emusPerInch);

                var ratWidth = wImgEmus / hImgEmus;
                long emuWidth = (long)(emuHeight * ratWidth);


                if (maxwidth != 0 && maxEmuWidth < emuWidth && scale == true)
                {
                    emuWidth = maxEmuWidth;
                    var ratHeight = hImgEmus / wImgEmus;
                    emuHeight = (long)(maxEmuWidth * ratHeight);
                }
                else if (maxwidth != 0 && maxEmuWidth < emuWidth && scale == false)
                {
                    emuWidth = maxEmuWidth;
                }

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuHeight, emuHeight));
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
            var SectionBreakProperties = new SectionProperties(new SectionType() { Val = SectionMarkValues.Continuous });

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
        public Table CreateNewTable(List<string[]> datosTabla, bool haveBorder = true, string subtitle = "")
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

                    SetCellStyles(ref texto, ref cProps);
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

                    SetCellStyles(ref texto, ref cProps);
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

        private static void SetCellStyles(ref string texto, ref TableCellProperties cProps)
        {
            var colorCell = texto.Contains("[CC:");

            if (colorCell)
            {
                var color = texto.Substring(texto.IndexOf("[CC:") + 4, 7);
                cProps.Append(new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = color });
                texto = texto.Replace("[CC:" + color + "]", "");
            }
        }

        private static void SetCellTextStyles(ref string texto, ref RunProperties rProps, ref ParagraphProperties pProps)
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
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "nn" };
                    break;

                case ParagraphTypes.Heading1:
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "tt1" };
                    break;

                case ParagraphTypes.Heading2:
                    paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "tt2" };
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


        public void CreateAndAddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename)
        {
            // Access the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;
            if (styles == null)
            {
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();
            }

            // Create a new paragraph style element and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            BasedOn basedon1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Ttulo1Car" };
            PrimaryStyle primaryStyle = new PrimaryStyle();


            ParagraphProperties pprops = new ParagraphProperties();
            NumberingProperties nprop = new NumberingProperties();
            nprop.Append(new NumberingId() { Val = 1 });
            nprop.Append(new NumberingLevelReference() { Val = 0 });

            pprops.Append(nprop);

            AutoRedefine autoredefine1 = new AutoRedefine() { Val = OnOffOnlyValues.Off };
            Locked locked1 = new Locked() { Val = OnOffOnlyValues.Off };
            PrimaryStyle primarystyle1 = new PrimaryStyle() { Val = OnOffOnlyValues.On };
            StyleHidden stylehidden1 = new StyleHidden() { Val = OnOffOnlyValues.Off };
            SemiHidden semihidden1 = new SemiHidden() { Val = OnOffOnlyValues.Off };
            StyleName styleName1 = new StyleName() { Val = stylename };
            UIPriority uipriority1 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhidewhenused1 = new UnhideWhenUsed() { Val = OnOffOnlyValues.On };


            style.Append(autoredefine1);
            style.Append(basedon1);
            style.Append(linkedStyle1);
            style.Append(locked1);
            style.Append(primarystyle1);
            style.Append(stylehidden1);
            style.Append(semihidden1);
            style.Append(styleName1);
            style.Append(nextParagraphStyle1);
            style.Append(uipriority1);
            style.Append(unhidewhenused1);
            style.Append(primaryStyle);
            style.Append(pprops);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "24" };


            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }








        public string GetTOC(string title, int titleFontSize)
        {
            return $@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
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
        <w:instrText xml:space='preserve'>TOC \h \z \t ""tt1,1,tt2,2""</w:instrText>
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
        }


        public void GetTOCTable(string title, int titleFontSize)
        {

        }
    }
}
