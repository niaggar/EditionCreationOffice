using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DIS.Reportes.Automatizados.Components
{
    public class CreateTable
    {
        /// <summary>
        /// Crear nueva tabla
        /// </summary>
        /// <param name="datosTabla"></param>
        /// <param name="haveBorder"></param>
        /// <returns></returns>
        public static Table NewT(List<string[]> datosTabla, bool haveBorder = true, bool rowHeader = true, int fontSize = 20)
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

                if (row == 0 && rowHeader)
                {
                    TableRowProperties tblHeaderRowProps = new TableRowProperties(
                        new CantSplit() { Val = OnOffOnlyValues.On },
                        new TableHeader() { Val = OnOffOnlyValues.On }
                    );
                    tableRow.AppendChild<TableRowProperties>(tblHeaderRowProps);
                }

                for (int col = 0; col < columnCount; col++)
                {
                    var texto = datosTabla[row].Length > col ? datosTabla[row][col] : "";

                    var pProps = new ParagraphProperties(
                        new SpacingBetweenLines() { Before = "40", After = "40" },
                        new Indentation() { Left = "40" },
                        new Languages() { Val = "es-ES" }
                    );
                    var rProps = new RunProperties(
                        new FontSize() { Val = fontSize.ToString() },
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

        /// <summary>
        /// Crear nueva tabla con imagenes en archivos
        /// </summary>
        /// <param name="datosTabla"></param>
        /// <param name="utilSpace"></param>
        /// <param name="mainPart"></param>
        /// <param name="haveBorder"></param>
        /// <returns></returns>
        public static Table NewTImg(List<string[]> datosTabla, (double width, double height) utilSpace, MainDocumentPart mainPart, bool haveBorder = true)
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
                        var img = CreateImages.NewImg(mainPart, texto, width: (long)(utilSpace.width / 2));
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

        /// <summary>
        /// Crear nueva tabla con imagenes en base 64
        /// </summary>
        /// <param name="datosTabla"></param>
        /// <param name="utilSpace"></param>
        /// <param name="mainPart"></param>
        /// <param name="haveBorder"></param>
        /// <returns></returns>
        public static Table NewTImg64(List<string[]> datosTabla, MainDocumentPart mainPart, (double width, double height) utilSpace, bool haveBorder = true)
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


                    var is64 = texto.Contains("[B64]");

                    if (is64)
                    {
                        texto = texto.Replace("[B64]", "");
                        // Get the with/height from the text, the standart is [WIDHT:2.0:END]
                        var haveWidth = texto.Contains("[WIDTH:");
                        var haveHeight = texto.Contains("[HEIGHT:");
                        double widht = 0, height = 0;
                        
                        if (haveWidth)
                        {
                            var start = texto.IndexOf("[WIDTH:") + 7;
                            var val = texto.Substring(start, 4);
                            widht = double.Parse(val);
                            texto = texto.Replace("[WIDTH:" + val + "]", "");
                        }
                        if (haveHeight)
                        {
                            var start = texto.IndexOf("[HEIGHT:") + 8;
                            var val = texto.Substring(start, 4);
                            height = double.Parse(val);
                            texto = texto.Replace("[HEIGHT:" + val + "]", "");
                        }

                        Paragraph img;
                        if (widht == 0 && height == 0)
                            img = new Paragraph();
                        else if (widht == 0 && height > 0)
                            img = CreateImages.NewImgB64Height(mainPart, texto, height, utilSpace.width);
                        else if (widht > 0 && height == 0)
                            img = CreateImages.NewImgB64Width(mainPart, texto, widht, utilSpace.height);
                        else
                            img = CreateImages.NewImgB64(mainPart, texto, width: widht, height: height);

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


        #region Utils
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
    }
}
