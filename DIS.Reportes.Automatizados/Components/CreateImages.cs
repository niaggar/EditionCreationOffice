using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Windows.Media.Imaging;
using System.Windows;

namespace DIS.Reportes.Automatizados.Components
{
    public static class CreateImages
    {
        private const int emusPerInch = 914400;


        #region Create new image element from a image file
        /// <summary>
        /// Crear nueva imagen de archivos
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="escale"></param>
        /// <returns></returns>
        public static Paragraph NewImg(MainDocumentPart mainPart, string fileName, double escale = 1)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

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

        /// <summary>
        /// Crear imagenes de archivos con width y height definidos
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImg(MainDocumentPart mainPart, string fileName, double width = 0, double height = 0)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

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

                long emuWidth = (long)(height * emusPerInch);
                long emuHeight = (long)(width * emusPerInch);

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuWidth, emuHeight));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        /// <summary>
        /// Crear imagenes de archivos con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImgWidth(MainDocumentPart mainPart, string fileName, double width, double maxheight, bool scale = true)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

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

                long emuWidth = (long)(width * emusPerInch);
                long maxEmuHeight = (long)(maxheight * emusPerInch);

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
        /// Crear imagenes de archivos con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImgHeight(MainDocumentPart mainPart, string fileName, double height, double maxwidth, bool scale = true)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

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

                long emuHeight = (long)(height * emusPerInch);
                long maxEmuWidth = (long)(maxwidth * emusPerInch);

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
        #endregion

        #region Create new image elemente from a base64 image
        /// <summary>
        /// Crear nueva imagen base 64
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="base64"></param>
        /// <param name="escale"></param>
        /// <returns></returns>
        public static Paragraph NewImgB64(MainDocumentPart mainPart, string base64, double escale = 1)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

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

        /// <summary>
        /// Crear imagenes base 64 con width y height definidos
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImgB64(MainDocumentPart mainPart, string base64, double width = 0, double height = 0)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

                byte[] imgBytes = Convert.FromBase64String(base64);
                BitmapImage img = new BitmapImage();

                img.BeginInit();
                img.StreamSource = new MemoryStream(imgBytes);
                img.EndInit();
                imagePart.FeedData(new MemoryStream(imgBytes));

                long emuWidth = (long)(width * 914400);
                long emuHeight = (long)(height * 914400);

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuWidth, emuHeight));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }

        /// <summary>
        /// Crear imagenes base 64 con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImgB64Width(MainDocumentPart mainPart, string base64, double width, double maxheight, bool scale = true)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

                byte[] imgBytes = Convert.FromBase64String(base64);
                BitmapImage img = new BitmapImage();

                img.BeginInit();
                img.StreamSource = new MemoryStream(imgBytes);
                img.EndInit();
                imagePart.FeedData(new MemoryStream(imgBytes));

                long emuWidth = (long)(width * emusPerInch);
                long maxEmuHeight = (long)(maxheight * emusPerInch);

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
        /// Crear imagenes base 64 con width definido y height maximo
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="fileName"></param>
        /// <param name="width">Pulgadas</param>
        /// <param name="height">Pulgadas</param>
        /// <returns></returns>
        public static Paragraph NewImgB64Height(MainDocumentPart mainPart, string base64, double height, double maxwidth, bool scale = true)
        {
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

                byte[] imgBytes = Convert.FromBase64String(base64);
                BitmapImage img = new BitmapImage();

                img.BeginInit();
                img.StreamSource = new MemoryStream(imgBytes);
                img.EndInit();
                imagePart.FeedData(new MemoryStream(imgBytes));

                long emuHeight = (long)(height * emusPerInch);
                long maxEmuWidth = (long)(maxwidth * emusPerInch);

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

                return CreateNewImageElement(mainPart.GetIdOfPart(imagePart), new Size(emuWidth, emuHeight));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return new Paragraph();
            }
        }
        #endregion


        /// <summary>
        /// Crea el elemento de OpenXML que permite agregar imagenes en un documento
        /// </summary>
        /// <param name="relationId"></param>
        /// <param name="size"></param>
        /// <returns></returns>
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
    }
}
