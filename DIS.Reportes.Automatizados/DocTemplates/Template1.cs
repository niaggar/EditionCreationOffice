using DIS.Reportes.Automatizados.Components;
using DIS.Reportes.Automatizados.Models.Enums;
using DIS.Reportes.Automatizados.Utils;
using sty = DIS.Reportes.Automatizados.Utils.StylesController;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace DIS.Reportes.Automatizados.DocTemplates
{
    public class Template1
    {
        private const string GRAY = "#F2F2F2";
        
        
        public static string GetSaveRoute()
        {
            var createFile = new SaveFileDialog()
            {
                FileName = "Template1.docx",
                Filter = "Word Files (*.docx)|*.docx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                DefaultExt = "docx"
            };
            var res = createFile.ShowDialog();
            if (res != true) return "";

            return createFile.FileName;
        }




        public static void CrearSubUperScript()
        {
            string route = GetSaveRoute();
            if (route == "") return;

            var fileDocument = DocumentController.Create(route);
            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());


            var OrdinalNumberSuffixRegex = new Regex("\\[(.*?)\\]");
            var innerText = "Esto es [^^prueba1^^] un texto e[^^2x^^] de prueba [__brueba2__]";
            var startIndex = 0;
            var destParagraph = new Paragraph();

            foreach (Match match in OrdinalNumberSuffixRegex.Matches(innerText))
            {
                if (match.Index > startIndex)
                {
                    string text = innerText.Substring(startIndex, match.Index - startIndex);
                    destParagraph.AppendChild(new Run(CreateText(text)));
                }

                VerticalTextAlignment aligment = new VerticalTextAlignment();

                if (match.Value.Contains("[^^"))
                {
                    aligment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };
                }
                else if (match.Value.Contains("[__"))
                    aligment = new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript };

                destParagraph.AppendChild(
                    new Run(
                        new RunProperties(
                            aligment,
                        CreateText(match.Value.Substring(3, match.Value.Length - 6)))));

                startIndex = match.Index + match.Length;
            }

            if (startIndex < innerText.Length)
            {
                string text = innerText.Substring(startIndex);
                destParagraph.AppendChild(new Run(CreateText(text)));
            }


            body.AppendChild(destParagraph);
            mainpart.Document.Save();
            fileDocument.Close();
        }


        public static Text CreateText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new Text();
            }

            if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
            {
                return new Text(text) { Space = SpaceProcessingModeValues.Preserve };
            }

            return new Text(text);
        }





        public static void Create()
        {
            string route = GetSaveRoute();
            if (route == "") return;


            // Configuraciones globales del documento
            var documentSize = PageSizeTypes.A4;
            var documentMargins = (top: 0.79, right: 0.98, bottom: 1.18, left: 1.38);


            var fileDocument = DocumentController.Create(route);
            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());


            // footer
            #region Crear Footer Global
            var globalFooterPart = mainpart.AddNewPart<FooterPart>();
            var globalFooterPartId = mainpart.GetIdOfPart(globalFooterPart);

            var name = System.IO.Path.GetFileNameWithoutExtension(route);
            var globalFooter = CreateUpdatableElement.NewFooter($"Archivo: {name}");
            globalFooter.Save(globalFooterPart);
            #endregion


            // estilos
            #region Estilos
            StyleDefinitionsPart part = doc.MainDocumentPart.StyleDefinitionsPart;
            if (part == null) part = sty.AddStylesPartToPackage(doc);
            sty.CreateAndAddParagraphStyle(part);
            sty.AddNumberingPart(doc);
            #endregion



            // Crear portada
            #region portada
            CreateFirstPage(ref mainpart, body);
            #endregion




            // Crear tabla de contenidos
            #region TOC
            CreateUpdatableElement.NewTOC(doc);
            body.AppendChild(CreateParagraph.NewPBreak());

            #region Crear Header del TOC
            var tocHeaderPart = mainpart.AddNewPart<HeaderPart>();
            var tocHeaderPartID = mainpart.GetIdOfPart(tocHeaderPart);

            var tocHeader = CreateUpdatableElement.NewHeader("MEMORIA DE DISEÑO DE ESTRUCTURAS METÁLICAS DE PÓRTICOS", "CO-RBAN: RENOVACIÓN SUBESTACIÓN BANADÍA 230 kV");
            tocHeader.Save(tocHeaderPart);
            #endregion

            #region Crear seccion
            // Crear seccion
            var pStoc = CreateSection.NewS();
            var stocProp = pStoc.Descendants<SectionProperties>().FirstOrDefault();

            // Agregar header y footer de seccion
            stocProp.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = tocHeaderPartID });
            stocProp.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });

            // Establecer tamaños
            SizeController.SetPageSize(stocProp, documentSize, PageOrientationValues.Portrait);
            SizeController.SetMarginSize(stocProp, documentMargins, PageOrientationValues.Portrait);
            #endregion

            body.AppendChild(pStoc);
            #endregion



            // Contenido
            #region Contenido
            #region Crear Header del contenido
            var contHeaderPart = mainpart.AddNewPart<HeaderPart>();
            var contHeaderPartID = mainpart.GetIdOfPart(contHeaderPart);

            var contHeader = CreateUpdatableElement.NewHeader("CRITERIOS DE DISEÑO ELECTROMECÁNICO", "CO-RBAN: RENOVACIÓN SUBESTACIÓN BANADÍA 230 KV");
            contHeader.Save(contHeaderPart);
            #endregion

            #region Crear seccion inicial
            // Crear seccion
            var pSection1 = CreateSection.NewS();
            var secProps1 = pSection1.Descendants<SectionProperties>().FirstOrDefault();

            // Agregar header y footer de seccion
            secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = contHeaderPartID });
            secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });

            // Establecer tamaños
            SizeController.SetPageSize(secProps1, documentSize, PageOrientationValues.Portrait);
            SizeController.SetMarginSize(secProps1, documentMargins, PageOrientationValues.Portrait);
            #endregion


            #region Contenido
            body.AppendChild(CreateParagraph.NewP("OBJETO", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("Presentar los procedimientos, criterios y resultados de los análisis efectuados para el diseño estructural de los pórticos metálicos requeridos para el cambio rápido del nuevo reactor de repuesto de 12.5 Mvar que será instalado en la subestación Banadía 230 kV, ubicada en el municipio de Saravena, en el departamento de Arauca.", ParagraphTypes.Normal));
            
            body.AppendChild(CreateParagraph.NewP("ALCANCE", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("En los siguientes capítulos se detallarán los procedimientos, criterios y resultados de los análisis efectuados para el diseño de la estructura metálica de los pórticos. Se incluye una descripción de las cargas aplicadas producto del peso de los equipos, cables, y de las acciones ambientales que inciden directamente sobre las estructuras metálicas. Además, se presentan los resultados del análisis y diseño realizado usando el software SAP2000, para cada uno de los elementos que conforman las estructuras atendiendo las solicitaciones más desfavorables que exijan las distintas combinaciones de carga.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Los diseños han sido realizados teniendo en cuenta todos los requerimientos de las especificaciones técnicas del proyecto [10] y [2]. Los resultados del diseño se ilustran en el plano “CO-RBAN-14113-S-01-K1525: Planos de diseño estructuras metálicas de pórticos”, en dicho plano se presenta la guía para la fabricación de las estructuras.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("DESCRIPCIÓN DE LOS PORTICOS", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("Los pórticos se diseñan como estructuras en celosía con diagonales, estos elementos soportan en la parte superior las cargas de templas y equipos dependiendo de la configuración del sistema. Además, los pórticos se encargan de transmitir las solicitaciones a la fundación y posteriormente al suelo.", ParagraphTypes.Normal));
            
            body.AppendChild(CreateParagraph.NewP("ESPECIFICACIONES DE LOS MATERIALES", ParagraphTypes.Heading1));
            body.AppendChild(CreateTable.NewT(DatosTabla1()));
            
            body.AppendChild(CreateParagraph.NewP("CRITERIOS DE DISEÑO", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("El diseño de la estructura metálica para los pórticos se lleva a cabo teniendo en cuenta los criterios de diseño de estructuras metálicas [10], documento en el que se referencian las especificaciones de los planos del fabricante de los equipos, la geometría básica, distancias eléctricas, y cargas de conexión.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("El análisis estructural se realizó en el software SAP2000, versión 24.0.0, mediante un modelo tridimensional, en el cual, la estructura está idealizada como un conjunto de celosías planas, con una configuración de diagonales tipo “X”.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Para el estado límite de resistencia, el diseño de los elementos se realizó con la aplicación IEB “Diseño de Estructura Metálica de pórticos y Equipos”, la cual con base en la información de entrada (resultados del SAP 2000), realiza el diseño por compresión, tracción, flexión, la interacción entre estas solicitaciones y el diseño de las conexiones para la cantidad mínima requerida de pernos.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("La determinación de los esfuerzos máximos a compresión, tensión, flexión, cortante y aplastamiento se hace siguiendo los lineamientos de las normas AISC 360 – 16 (American Institute of Steel Construction), referencia [11], y ASCE 10-15 (American Society of Civil Engineers) “Design of Latticed Steel Transmission Structures” referencia [12] y siguiendo las recomendaciones del manual ASCE N°52 “Guide for Design of Steel Transmission Towers”, referencia [13]; con ayuda del programa SAP2000.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Para la definición de los elementos metálicos los límites de las relaciones de esbeltez serán los presentados en la Tabla 2:", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla2()));
            body.AppendChild(CreateParagraph.NewP("La dimensión mínima de los perfiles que componen las estructuras debe responder a la Tabla 3.", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla3()));

            body.AppendChild(CreateParagraph.NewP("CRITERIOS DE DEFLEXIONES", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("Las deformaciones de la estructura metálica, se limitarán a los valores presentados en la Tabla 4 (Los valores fueron tomados del capítulo 4 de la norma ASCE 113 “Substation Structure Design Guide” referencia [4])", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla4()));
            body.AppendChild(CreateParagraph.NewP("Los elementos de las estructuras de pórticos se deben clasificar como tipo A, cuando hay equipos sobre los pórticos y tipo B cuando no se tienen equipos sobre los pórticos. ", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Los elementos de las estructuras de soporte de los seccionadores e interruptores se deben clasificar como tipo A y como tipo B las estructuras de soporte de los demás equipos.", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla5(), rowHeader: false));
            body.AppendChild(CreateParagraph.NewP($"Notas:{sty.SetBold()}{sty.SetFontSize(24)}{sty.SetUnderline()}{sty.SetLeftAligment()}", ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP("La luz para los miembros horizontales debe ser medida como la luz libre entre miembros verticales o para miembros en cantiléver como la distancia al punto vertical más cercano. Luego la deflexión debe ser el desplazamiento neto, vertical u horizontal, relativo al punto de soporte.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("La luz para miembros verticales debe ser la distancia vertical desde el punto de conexión de la fundación al punto de investigación.", ParagraphTypes.Normal));

            body.AppendChild(CreateParagraph.NewP("CARGAS", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("Para el diseño de la estructura se considera el peso propio entre las cargas actuantes. En las cargas de diseño presentadas en los planos no se incluyen factores de sobrecarga, por lo tanto, en el análisis de la estructura metálica realizado se incluyen estos factores.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Las cargas sobre los pórticos y las dimensiones generales son tomadas de los documentos de referencia del [20] al [22].", ParagraphTypes.Normal));
            
            body.AppendChild(CreateParagraph.NewP("PESO PROPIO DE LA ESTRUCTURA", ParagraphTypes.Heading2));
            body.AppendChild(CreateParagraph.NewP("Cargas debidas al peso de la estructura metálica, cables, templas, aisladores, herrajes, accesorios, y todos los elementos que componen el conjunto analizado. Se afecta en un 20% adicional para considerar el peso de los elementos estructural no modelados tales como: platinas, pernos, tuercas, arandelas, galvanizado, etc.", ParagraphTypes.Normal));
            
            body.AppendChild(CreateParagraph.NewP("CARGAS DE CONEXIÓN", ParagraphTypes.Heading2));
            body.AppendChild(CreateParagraph.NewP("Se refiere a las tensiones mecánicas y cargas de cortocircuito. Considerando las tensiones mecánicas, esta es aplicable a barraje flexible en templas, barras, cable guardas, conexión entre equipos, etc. Metodología según Overhead Power Lines, referencia [15]. Flecha máxima para condición EDS del barraje del 3%.", ParagraphTypes.Normal));
            
            body.AppendChild(CreateParagraph.NewP("CARGAS DE VIENTO", ParagraphTypes.Heading2));
            body.AppendChild(CreateParagraph.NewP("Se considera las cargas de vientos sobre templas, equipos y estructuras en dirección X y Y. La velocidad del viento se toma de la NSR-10 [2] y el cálculo de estas fuerzas se realiza bajo la metodología del manual ASCE-74 “Guidelines for Electrical Transmission Line Structural Loading”, referencia [3]. La fuerza del viento sobre la estructura debida a la presión del viento sobre los conductores se calcula como:", ParagraphTypes.Normal));
            var eq1 = "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";
            var eq2 = "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";
            body.AppendChild(CreateImages.NewImgB64Height(mainpart, eq1, 0.25, 999));
            body.AppendChild(CreateImages.NewImgB64Height(mainpart, eq2, 0.25, 999));
            body.AppendChild(CreateParagraph.NewP("Donde:", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTablaDonde1(), rowHeader: false, haveBorder: false, fontSize: 24));

            body.AppendChild(CreateParagraph.NewP("CARGAS DE SISMO", ParagraphTypes.Heading2));
            body.AppendChild(CreateParagraph.NewP("El cálculo de estas fuerzas se realiza bajo la metodología del Reglamento Colombiano de Construcción Sismo Resistente NSR-10, referencia [2]. El sismo vertical se define como Ez = 2/3 E(x,y). Los parámetros sísmicos se indican en la siguiente referencia [9].", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Donde:", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Ez: \tSismo vertical", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Ex,y: Sismo horizontal", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP($"Nota:{sty.SetBold()}{sty.SetFontSize(24)}{sty.SetUnderline()}{sty.SetLeftAligment()}", ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP("para el análisis estructural se utilizó coeficiente de capacidad de disipación de energía R de 3.00 y un factor de sobre-resistencia Ω de 3.00.", ParagraphTypes.Normal));

            body.AppendChild(CreateParagraph.NewP("CARGAS DE MONTAJE Y MANTENIMIENTO", ParagraphTypes.Heading2));
            body.AppendChild(CreateParagraph.NewP("Todos los miembros de las estructuras en análisis cuyo eje longitudinal forme un ángulo con la horizontal menor que 45 grados tendrán suficiente sección para resistir una carga adicional de 150 daN vertical, aplicada en cualquier punto de su eje longitudinal.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Considerando las cargas de montaje y mantenimiento para columnas: el castillete será diseñado para resistir la acción de un hombre con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 150 daN.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Considerando las cargas de montaje y mantenimiento para vigas: el nodo donde llega cada barraje, será diseñado para resistir la acción de dos hombres con herramienta de montaje que equivale a aplicar verticalmente un peso aproximado de 250 daN.", ParagraphTypes.Normal));

            body.AppendChild(CreateParagraph.NewP("COMBINACIONES DE CARGA", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("Para el diseño de la estructura metálica se utilizan las combinaciones de carga que se listan de la Tabla 6 a la Tabla 8; estas combinaciones de carga provienen del documento CO-RBAN-14113-S-01-D1181 “Criterios de diseño - Estructuras metálicas” [10].", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla6(), rowHeader: true));
            body.AppendChild(CreateParagraph.NewP("", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla7(), rowHeader: true));
            body.AppendChild(CreateParagraph.NewP("", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTabla8(), rowHeader: true));
            body.AppendChild(CreateParagraph.NewP($"Nota:{sty.SetBold()}{sty.SetFontSize(24)}{sty.SetUnderline()}{sty.SetLeftAligment()}", ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP("La fuerza de peso propio y Ez se toman positivas en el sentido de la gravedad.", ParagraphTypes.Normal));
            body.AppendChild(CreateParagraph.NewP("Donde:", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTablaDonde2(), rowHeader: false, haveBorder: false, fontSize: 24));
            body.AppendChild(CreateParagraph.NewP($"Nota:{sty.SetBold()}{sty.SetFontSize(24)}{sty.SetUnderline()}{sty.SetLeftAligment()}", ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP("La fuerza de peso propio y Ez se toman positivas en el sentido de la gravedad.", ParagraphTypes.Normal));

            body.AppendChild(CreateParagraph.NewP("NOMENCLATURA DEL REPORTE", ParagraphTypes.Heading1));
            body.AppendChild(CreateParagraph.NewP("A continuación, se indica la nomenclatura del reporte del diseño de ángulos del soporte crítico que será presentado posteriormente.", ParagraphTypes.Normal));
            body.AppendChild(CreateTable.NewT(DatosTablaNomenclatura(), rowHeader: true, haveBorder: false, fontSize: 24));




            #endregion


            body.AppendChild(pSection1);
            #endregion



            // Final
            #region Seccion final
            #region Crear final seccion
            var secFinal = CreateSection.NewSFinal();

            SizeController.SetPageSize(secFinal, documentSize, PageOrientationValues.Portrait);
            SizeController.SetMarginSize(secFinal, documentMargins, PageOrientationValues.Portrait);
            #endregion

            // Agregando seccion final
            body.AppendChild(secFinal);
            #endregion



            mainpart.Document.Save();
            fileDocument.Close();

            DocumentController.SavePDF(route, route + ".pdf");
        }




        public static List<string[]> DatosTabla1()
        {
            var datos = new List<string[]>
            {
                new string[3]
            {
                $"ÍTEM{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"DESCRIPCIÓN{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"CRITERIO{sty.SetCellColor(GRAY)}{sty.SetBold()}"
            },
                new string[3]
            {
                "Elemento", $"Perfiles{sty.SetLeftAligment()}", $"ASTM A-572 Gr50{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Platinas{sty.SetLeftAligment()}", $"ASTM A-36{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Soldadura{sty.SetLeftAligment()}", $"De acuerdo AWS D1.1 y D1.3.\r\nElectrodos E70-XX{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Tornillos{sty.SetLeftAligment()}", $"ASTM A-394 TIPO 0{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Pernos de anclaje{sty.SetLeftAligment()}", $"F1554 Gr 55{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Arandelas{sty.SetLeftAligment()}", $"ASTM F-436{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Tuercas{sty.SetLeftAligment()}", $"ASTM A-563{sty.SetLeftAligment()}"
            },
                new string[3]
            {
                "|", $"Galvanización{sty.SetLeftAligment()}", $"ASTM A-123, ASTM A-153{sty.SetLeftAligment()}"
            }
            };


            return datos;
        }

        public static List<string[]> DatosTabla2()
        {
            var datos = new List<string[]>
            {
                new string[3] { $"ÍTEM{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"DESCRIPCIÓN{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"CRITERIO{sty.SetCellColor(GRAY)}{sty.SetBold()}" },
                new string[3] { $"Relación de esbeltez - ASCE 10-15", $"Otros miembros", $"L/r ≤ 200" },
                new string[3] { $"|", $"Redundantes", $"L/r ≤ 250" },
                new string[3] { $"|", $"Solo a tensión", $"L/r ≤ 350" },
                new string[3] { $"|", $"Miembros a compresión", $"Montantes L/r ≤ 150" },
                new string[3] { $"Relación w/t - ASCE 10-15", $"Ángulos a 90° Numeral 3.7.1", $"Máximo w/t ≤ 25" },
                new string[3] { $"|", $"Compacto", $"w/t ≤ (w/t) lím" },
                new string[3] { $"|", $"Esbelto Ecuación 3.7-2", $"(w/t) lím< w/t ≤144Ψ/Fy^1/2" },
                new string[3] { $"|", $"Esbelto Ecuación 3.7-3", $"w/t >144Ψ/Fy^1/2" },
            };

            return datos;
        }

        public static List<string[]> DatosTabla3()
        {
            var datos = new List<string[]>
            {
                new string[3] { $"ÍTEM{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"DESCRIPCIÓN{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"CRITERIO{sty.SetCellColor(GRAY)}{sty.SetBold()}" },
                new string[3] { $"Espesor mínimo - ASCE 10-15", $"Miembros", $"3/16\" (4.8mm)" },
                new string[3] { $"|", $"Miembros secundarios redundantes", $"1/8\" (3.2mm)" },
                new string[3] { $"|", $"Platinas de conexión", $"3/16\" (4.8mm)" },
                new string[3] { $"|", $"Criterio de espesor exposición a corrosión", $"3/16\" (4.8mm)" },
            };

            return datos;
        }

        public static List<string[]> DatosTabla4()
        {
            var datos = new List<string[]>
            {
                new string[5] { $"TIPO DE DEFLEXIÓN{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"ESTRUCTURAS CLASE A{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"~", $"ESTRUCTURAS CLASE B{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"~" },
                new string[5] { $"|", $"Elementos horizontales{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"Elementos verticales{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"Elementos horizontales{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"Elementos verticales{sty.SetCellColor(GRAY)}{sty.SetBold()}" },
                new string[5] { $"Horizontal", $"1/200", $"1/100", $"1/100", $"1/100" },
                new string[5] { $"Vertical", $"1/200", $"", $"1/200", $"" },
            };

            return datos;
        }

        public static List<string[]> DatosTabla5()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"Clase A", $"Interruptores y seccionadores", },
                new string[2] { $"Clase B", $"Transformadores de corriente, transformadores de tensión, descargadores de sobretensión, aisladores poste y trampas de onda", },
                new string[2] { $"|", $"Columnas de pórticos - Vigas de pórticos", },
            };

            return datos;
        }

        public static List<string[]> DatosTabla6()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"COMBINACIONES DE CARGA{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"~", },
                new string[2] { $"Diseño Estructural{sty.SetLeftAligment()}", $"1) 1,2 PP + 1,3 CT+ 1,0 CMM{sty.SetLeftAligment()}", },
                new string[2] { $"|", $"2) 1,1 PP + 1,1 CT ± 1,0 VD(X,Y) + 1,0 CTVDL{sty.SetLeftAligment()}", },
                new string[2] { $"|", $"3) 1,1 PP + 1,1 CT ± 1,0 E(X,Y) ± 0,3 E(Y,X) + 1,0 E(Z){sty.SetLeftAligment()}", },
                new string[2] { $"|", $"4) 0,9 PP + 1,1 CT ± 1,0 E(X,Y) ± 0,3 E(Y,X) - 1,0 E(Z){sty.SetLeftAligment()}", },
                new string[2] { $"|", $"5) 1,1 PP + 1,1 CT + 1,0 CC + 1,0 AC{sty.SetLeftAligment()}", },
            };

            return datos;
        }

        public static List<string[]> DatosTabla7()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"COMBINACIONES DE CARGA{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"~", },
                new string[2] { $"Estabilidad{sty.SetLeftAligment()}", $"1) 1,0 PP + 1,0 CT + 1,0 CMM{sty.SetLeftAligment()}", },
                new string[2] { $"|", $"2) 1,0 PP + 1,0 CT ± 1,0 VS(X,Y) + 1,0 CTVSL{sty.SetLeftAligment()}", },
                new string[2] { $"|", $"3) 1,0 PP + 1,0 CT ± 0,7 E(X,Y) ± 0,21 E(Y,X) + 0,7 E(Z){sty.SetLeftAligment()}", },
                new string[2] { $"|", $"4) 0,6 PP + 1,0 CT ± 0,7 E(X,Y) ± 0,21 E(Y,X) - 0,7 E(Z){sty.SetLeftAligment()}", },
                new string[2] { $"|", $"5) 1,0 PP + 1,0 CT + 1,0 CC + 1,0 AC{sty.SetLeftAligment()}", },
            };

            return datos;
        }

        public static List<string[]> DatosTabla8()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"Condición{sty.SetCellColor(GRAY)}{sty.SetBold()}", $"Combinación{sty.SetCellColor(GRAY)}{sty.SetBold()}", },
                new string[2] { $"Viento máximo esperado{sty.SetLeftAligment()}", $"1) 1.0PP + 0.78 V + 1.0 CT{sty.SetLeftAligment()}", },
                new string[2] { $"Accionamiento de equipos{sty.SetLeftAligment()}", $"2) 1.0PP + AC + 1.0 CT{sty.SetLeftAligment()}", },
            };

            return datos;
        }




        public static List<string[]> DatosTablaDonde1()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"F{sty.SetLeftAligment()}", $"Fuerza debida al viento{sty.SetLeftAligment()}", },
                new string[2] { $"γW{sty.SetLeftAligment()}", $"Factor de carga, función del periodo de retorno, Tabla 1-1 ó 1-2 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"V50{sty.SetLeftAligment()}", $"Velocidad del viento, para un periodo de retorno de 50 años. {sty.SetLeftAligment()}", },
                new string[2] { $"A{sty.SetLeftAligment()}", $"Área frontal efectiva de la estructura, en m^2.{sty.SetLeftAligment()}", },
                new string[2] { $"KZ{sty.SetLeftAligment()}", $"Coeficiente de exposición, Tabla 2-2 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"KZT{sty.SetLeftAligment()}", $"Factor de Topografía, Ec. 2-14 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"Q{sty.SetLeftAligment()}", $"Constante numérica, en función de la densidad del aire, sección 2.1.2 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"G{sty.SetLeftAligment()}", $"Factor de ráfaga. Sección 2.1.5 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"CF{sty.SetLeftAligment()}", $"Coeficiente de fuerza, sección 2.1.6 de referencia [3]{sty.SetLeftAligment()}", },
                new string[2] { $"qZ{sty.SetLeftAligment()}", $"Presión del viento{sty.SetLeftAligment()}", },
            };

            return datos;
        }

        public static List<string[]> DatosTablaDonde2()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"PP{sty.SetLeftAligment()}", $"Peso propio de la estructura (Ps), equipos (Pe) y conductores de la conexión (Pc).{sty.SetLeftAligment()}", },
                new string[2] { $"CT{sty.SetLeftAligment()}", $"Cargas por tensión mecánica de los conductores de conexión y cables guarda, se debe considerar tiro unilateral (un solo sentido, caso más desfavorable).{sty.SetLeftAligment()}", },
                new string[2] { $"CMM{sty.SetLeftAligment()}", $"Carga de montaje y mantenimiento.{sty.SetLeftAligment()}", },
                new string[2] { $"VD{sty.SetLeftAligment()}", $"Carga viento de diseño sobre equipos, cables y estructuras. Ver velocidad de viento de diseño en referencia [2]. Está conformado por el Viento sobre la estructura (VSx,y), sobre equipos (VEx,y) y sobre conductores de la conexión (VCx,y).{sty.SetLeftAligment()}", },
                new string[2] { $"VS{sty.SetLeftAligment()}", $"Carga viento de servicio sobre equipos, cables y estructuras. Ver velocidad de viento de servicio en referencia [2].{sty.SetLeftAligment()}", },
                new string[2] { $"CTVDL o CTVSL{sty.SetLeftAligment()}", $"Carga de sobretensión en el cable debido al viento de diseño o viento de servicio (solo actúa en el sentido de la tensión del cable). Correspondiente a (VCTx,y).{sty.SetLeftAligment()}", },
                new string[2] { $"CC{sty.SetLeftAligment()}", $"Cargas sobre conductores por efecto de cortocircuito.{sty.SetLeftAligment()}", },
                new string[2] { $"EX,Y{sty.SetLeftAligment()}", $"Cargas por sismo horizontal sobre equipos y estructuras, obtenidos con coeficientes sísmicos últimos. Para verificación de deflexiones, la carga sísmica se debe emplear sin dividir por R.{sty.SetLeftAligment()}", },
                new string[2] { $"EZ{sty.SetLeftAligment()}", $"Cargas por sismo vertical sobre equipos y estructuras, obtenidos con coeficientes sísmicos últimos.{sty.SetLeftAligment()}", },
                new string[2] { $"AC{sty.SetLeftAligment()}", $"Carga de accionamiento que aplica solo para interruptores.{sty.SetLeftAligment()}", },
            };

            return datos;
        }



        public static List<string[]> DatosTablaNomenclatura()
        {
            var datos = new List<string[]>
            {
                new string[2] { $"NOMENCLATURA DE REPORTE (1){sty.SetBold()}", $"~", },
                new string[2] { $"Elemento{sty.SetLeftAligment()}{sty.SetBold()}", $"Identificación del elemento en el software{sty.SetLeftAligment()}", },
                new string[2] { $"L:{sty.SetLeftAligment()}{sty.SetBold()}", $"Longitud no arriostrada del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Rx,y:{sty.SetLeftAligment()}{sty.SetBold()}", $"Radio de giro del elemento respecto a los ejes geométricos X y Y{sty.SetLeftAligment()}", },
                new string[2] { $"Ru:{sty.SetLeftAligment()}{sty.SetBold()}", $"Radio de giro del elemento respecto al eje principal menor U{sty.SetLeftAligment()}", },
                new string[2] { $"L/R:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de esbeltez{sty.SetLeftAligment()}", },
                new string[2] { $"kL/R:{sty.SetLeftAligment()}{sty.SetBold()}", $"Longitud efectiva{sty.SetLeftAligment()}", },
                new string[2] { $"Curva:{sty.SetLeftAligment()}{sty.SetBold()}", $"Ecuación empleada para la estimación de la longitud efectiva{sty.SetLeftAligment()}", },
                new string[2] { $"(L/R)LIM{sty.SetLeftAligment()}{sty.SetBold()}", $"Máxima relación de esbeltez{sty.SetLeftAligment()}", },
                new string[2] { $"Cc:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de esbeltez critica{sty.SetLeftAligment()}", },
                new string[2] { $"Fa:{sty.SetLeftAligment()}{sty.SetBold()}", $"Esfuerzo de compresión{sty.SetLeftAligment()}", },
                new string[2] { $"Comb:{sty.SetLeftAligment()}{sty.SetBold()}", $"Combinación de carga empleada en el diseño{sty.SetLeftAligment()}", },
                new string[2] { $"P{sty.SetLeftAligment()}{sty.SetBold()}", $"Fuerza axial de tracción o compresión*{sty.SetLeftAligment()}", },
                new string[2] { $"Puc:{sty.SetLeftAligment()}{sty.SetBold()}", $"Fuerza axial de compresión{sty.SetLeftAligment()}", },
                new string[2] { $"Put:{sty.SetLeftAligment()}{sty.SetBold()}", $"Fuerza axial de tracción{sty.SetLeftAligment()}", },
                new string[2] { $"V2:{sty.SetLeftAligment()}{sty.SetBold()}", $"Fuerza cortante en el plano 1-2{sty.SetLeftAligment()}", },
                new string[2] { $"V3:{sty.SetLeftAligment()}{sty.SetBold()}", $"Fuerza cortante en el plano 1-3{sty.SetLeftAligment()}", },
                new string[2] { $"M2:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento flector en el plano 1-3 (alrededor del eje 2){sty.SetLeftAligment()}", },
                new string[2] { $"M3:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento flector en el plano 1-2 (alrededor del eje 3){sty.SetLeftAligment()}", },
                new string[2] { $"Mr:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento actuante resultante, debido a M2 y M3{sty.SetLeftAligment()}", },
                new string[2] { $"θ°:{sty.SetLeftAligment()}{sty.SetBold()}", $"Angulo del momento resultante con respecto a la horizontal{sty.SetLeftAligment()}", },
                new string[2] { $"Uso:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de uso total del elemento (interacción de todas las solicitaciones){sty.SetLeftAligment()}", },
                new string[2] { $"Ecu:{sty.SetLeftAligment()}{sty.SetBold()}", $"Ecuación empleada para estimar el uso{sty.SetLeftAligment()}", },
                new string[2] { $"Puc/Pac:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de uso del elemento en compresión{sty.SetLeftAligment()}", },
                new string[2] { $"Put/Pat-v:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de uso del elemento en tracción{sty.SetLeftAligment()}", },
                new string[2] { $"Mr/Ma:{sty.SetLeftAligment()}{sty.SetBold()}", $"Relación de uso del elemento en flexión{sty.SetLeftAligment()}", },
                new string[2] { $"Pac:{sty.SetLeftAligment()}{sty.SetBold()}", $"Capacidad a compresión del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Pat-g:{sty.SetLeftAligment()}{sty.SetBold()}", $"Capacidad a tracción en el área bruta del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Pat-v:{sty.SetLeftAligment()}{sty.SetBold()}", $"Capacidad a tracción en el área neta del elemento o por bloque de cortante{sty.SetLeftAligment()}", },
                new string[2] { $"Pat:{sty.SetLeftAligment()}{sty.SetBold()}", $"Capacidad a tracción del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Ma:{sty.SetLeftAligment()}{sty.SetBold()}", $"Capacidad a flexión del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Pe:{sty.SetLeftAligment()}{sty.SetBold()}", $"Carga critica de pandeo de Euler{sty.SetLeftAligment()}", },
                new string[2] { $"ØPyc:{sty.SetLeftAligment()}{sty.SetBold()}", $"Resistencia axial del elemento en el área bruta{sty.SetLeftAligment()}", },
                new string[2] { $"Myt:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento que produce esfuerzos de tracción en la fibra extrema{sty.SetLeftAligment()}", },
                new string[2] { $"Myc:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento que produce compresión en la fibra extrema{sty.SetLeftAligment()}", },
                new string[2] { $"Me:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento crítico elástico {sty.SetLeftAligment()}", },
                new string[2] { $"Me.Ecu:{sty.SetLeftAligment()}{sty.SetBold()}", $"Ecuación empleada para el cálculo de Me, {sty.SetLeftAligment()}", },
                new string[2] { $"Mb:{sty.SetLeftAligment()}{sty.SetBold()}", $"Momento que produce pandeo lateral{sty.SetLeftAligment()}", },
                new string[2] { $"Mb.Ecu:{sty.SetLeftAligment()}{sty.SetBold()}", $"Ecuación empleada para el cálculo de Mb,{sty.SetLeftAligment()}", },
                new string[2] { $"K:{sty.SetLeftAligment()}{sty.SetBold()}", $"Factor que depende de la condición de apoyo del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"Cm:{sty.SetLeftAligment()}{sty.SetBold()}", $"Factor que depende de la distribución de momento en la sección{sty.SetLeftAligment()}", },
                new string[2] { $"Ø:{sty.SetLeftAligment()}{sty.SetBold()}", $"Diámetro del perno en pulgadas{sty.SetLeftAligment()}", },
                new string[2] { $"emin:{sty.SetLeftAligment()}{sty.SetBold()}", $"Distancia mínima al borde del elemento cortado{sty.SetLeftAligment()}", },
                new string[2] { $"fmin:{sty.SetLeftAligment()}{sty.SetBold()}", $"Distancia mínima al borde del elemento{sty.SetLeftAligment()}", },
                new string[2] { $"smin:{sty.SetLeftAligment()}{sty.SetBold()}", $"Distancia mínima entre centros de perforaciones{sty.SetLeftAligment()}", },
                new string[2] { $"(1)\tVer capítulos 3 y 4 del ASCE 10-15 [12]", $"~", },
            };

            return datos;
        }




        private static void CreateFirstPage(ref MainDocumentPart mainpart, Body body)
        {
            var margin = (top: 0.79, right: 0.69, bottom: 0.59, left: 0.98);

            #region GetData
            var title1 = "RENOVACIÓN DE SUBESTACIÓN";
            var title2 = "INGENIERÍA DE DETALLE PARA EL MONTAJE DE UN REACTOR DE REPUESTO 12,5 Mvar EN LA SUBESTACIÓN BANADÍA 230 kV";
            var title3 = "MEMORIA DE DISEÑO DE ESTRUCTURAS METÁLICAS DE PÓRTICOS";

            
            var tableAprobacion = new List<string[]>()
            {
                new string[7]{ "", "", "", "", "", "", "" },
                new string[7]{ "", "", "", "", "", "", "" },
                new string[7]{ "", "", "", "", "", "", "" },
                new string[7]{ "", "", "", "", "", "", "" },
                new string[7]{ "Estado/fase" + sty.SetBold(), "Rev." + sty.SetBold(), "Comentarios / Modificaciones" + sty.SetBold(), "Fecha de Act." + sty.SetBold(), "Elaboró" + sty.SetBold(), "Revisó" + sty.SetBold(), "Aprobó" + sty.SetBold() },
            };


            var imagenIEB = "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";
            var imagenCliente = "iVBORw0KGgoAAAANSUhEUgAAAGcAAABrCAYAAABqg5yCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAA3SSURBVHhe7ZwJdFRFFoYbBxzUWdwR2SGyJ+wKDiKiGDdAWQZUQJCBAzijg8igg4woogJyXGBUEHcQHccNCMeFTRZBFFCGLelOJ530kq2zdafTnXT3P/dWv0hCHsl7naYp5f3n3ENO+lWlqr5Xt27dqsYUDof3GyanmWBIWhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCSWAUdiGXAklgFHYhlwJJYBR2IZcCRWTOGEFTMUG8UETvnen+Be+iZyJswTVrB4NXx7flQ+PXMKHLeg9LNN8KzfQra5tn0e+de38weEyvxKKXnUIDiVOQXImbEQ5iY9kGZqD7Opi7BUUweYf5OInGkLUJnrVp6Ov/L/tQxHTL+j9iSINqnZcVNbansn2Ic9gAp7rlJSDkUNJ1hYDPvw6dTBlgSkOyymnjWMf5dqagX70GkIFpUopeIr91MraPBbUHt61WpfTUui51oh75FlSkk5FBWccDBIb+Vy6lBL0TH1DrNxp1sjfw51OhRSSsdP7oVa4fDL1AVZiaMRrqhUSp95RQWnwuZE+oUDqEPdVDta3dJMnWG9bDAC/zMrpeMnvXBs7W5D2F+hlD7zigqO5+OvaNA7Uqd61Oqkmh03tUHphylK6fhJN5wOt//y4RQteUuZNdrgpNGiW/jSu0rp+OnshLP8fRrwLtQprXBao/i1D5TS8dNZCce3fR91phN1SgscDrM7wrvxG6W0PoUrgwjmFaIiPRt+Wrf8Px4XFjhsRkWGHaFij/JkbZ2VcIKFpbD1HEODfpVqR6tbmqkdbP3GIZijfb/jP3AcJav+i/x/LoPjzr8ho+twWC7oR6F5Ag12a1Enh+rWZkNgv/5+5M1ZjNL3NtC+q+bfOCvhsEo++IIGqRt1qjN1Tq3zvejzzkg7txtK121SSqkrXFYO/yEz3ItWwT5iKrL6jEFW0mhk3zQZrsnzaPCfR8H85cJy/74ErrH/QFbPcbCc31cAO2a6jP5tSy/MCPFMML9I1Ote9MrZCQfhMAqfo8DgvB40MG2oc4nVOppIYGjn3bQH3M+uVgrUVsjrQ9m2fch/fAWcY+fC/czrKNu6J+KuikppzxFUnqwp4epKvDQb8+Fdvw2u8Y/B2mYIQWpGMNois9sIlO89BPfi16ht+uBUV6W7GIG0TJrJx1D+3SGU7zuMwBELKl35cckhRg9HkWfDNjjumglLU3Y7bYVZmvaFY/h0eD7bqjxVW4Ej6ShZk4LC5etE5xva2YAliyAvR/qlfyJAzWG9dBBsvUeJQdcKJ6vTCFRkOlD06ofImbkA2YMmIaNVMtLPu4Y+p/X1nF5U72Bk970brkmPonDJG/B9f1hpQezVYDgsdgXl1Ejvtj305n8rfg6VB5RPa6vCmQffD0dpIJzKb2Kn8oPHkZ38FzGbOWdmMfWuBULdeiD9/P7I6DyMynF+kNe2jgSlKxl7Bc6EJNHPEVfO9aea2sN6xfVw3D4Dnk07EArE1iXGBI4ehckdhisrT6tbCHrK4JrwGA0uByzawv2IJSlAGUR9s40/70Gguoi/Y27cHY4RDwgvECtphsM5J9/WfeTjt6Ns4456jZ/z7TyA0BlaYEO+cjhHP0xvdwcxiOoDHDvj2cXBieXSa1C44n1ekhsszXCChSWwdRtFEVIfpF8woH4jF2G5uD+yrh0H3/7T55frEi/otl4jhXtSG9DYmxKh0uzLnbGwTteuRdrhFBQjo8VQ+sMJit/VYl1pcb4S9lumIVTqVWrSJx7gIEVu0cq7aSfM5/DAnYgmT6+xu+PjkgTkTJyHELnYaKUDThEy2iRH/KuIgNQaVtvEWc95V6P8W+0no8EiD9zPrYZ95Ayaefcie+B4CrVnw5OyQ3lCu9i7OO+aRe2of8McW+PtRALy570caUgU0r7m+AMUNm9H4YtrRHaAZ4V6o042jnC6ouzL3UpNdcufmiFcIb8EkYgrQRhHT+YLeqPopXW0oOhz6F5qd1ojrcEBp5v4b7cm460BR2WRyE3v2iW8R+NEsd2IRprhVJf/4DHay3CIqsVVMJzuKPvqW6X0qcUbz6zrJpIrbC3KnRwxcT7PfFE/+PYdUkpoU4XNBVuXO6k8ZzOqt62m8WCmUVuzr52AnBkLkP/w88iftQQ5U+cjqy+/MBxi64PEoDMTRyLo1n8aHBWcyiwXrBcOFIOu1qCaph1OyXvrqfMcXanVw9aLPm9Hi+3TCOs4WQ3TTHNOfFTMAPV6I2Csl18Pz8btP6d/qquS9mae9dtgv3mqmEXa+s71JiG1UUcULn1bqUm7ooLDGWLr76+NKRze/7hGPUzPnvrt5l06D4z1kuvEHQY9yp+/QoA91Vsv0jdXDVOePrVCPj8Knl2JNNrX8MKvVtfJxpvV7AETdN+lkAcO7a4zWnPA0YGe51155CYPw+A1J3JLpjvSLxkIW9/RqHDlKSW1qeS9jbA07k11qrtiAUdH4tO95E1qj9ZjExqDc7vTPvE7pbQ2yQPH70dGl1thvrwfMukNdgyZipxJ81Gw4BWUvP0pvFt2U8R3UNxFqMjK0X0Rw7frAKwXX0dtUQ9k9MLhTaZz+EPixVGr72TjwKLwhXeU0tokD5xgEAGrHcHiUoTKysXgxzLFE8h0wNpmKLVF3W3qhcPyUmhvbsozp/7AiD1Azn2Pi35qlTRwTreCJR5kdryF2sKuqHY7o4FTSWtI1jX3aJo9XD8fDIa82jelvwo4nA3m9FJldg4qzDb49hxEydoUuJ9+Hbn3P0GDMgWZSSMo/O9D7akjINAJh+UcN0u4LLU6qxu7U1vSGITIM2jVLwZOiDbBvNaU7z8q3EnJu+tRuOJdcWHRNW6uiIb42Np8ThJtOLuK4CGSQuKgopNwKxEw6tnmaOHkjp+rA87oXy4cngG8Nvh27EfJmo1wL1xJG8AnxdtpHzoFth5jkNH8JnHoxWEx3zjl+wSRfQdHd3xd68TZSwSGlmgqejiu8XM0w8keOElXrk0qOJ6NW2Fu1gvpfxwAS5M+okO8R4ikT9orEDjaqlqA6ztz0W5RwaGQzXEH3xfn/ZN6vVXGgYjrnkfFEbtWSQWnfO+P1FHOevMs4Lr57VerU48xyPpnTzRw/EcsyGh5syirVmd14zyde/Gp71OoSSo4wWIv7IOniNnS0FnBAyZmWmOqhzafas9Ut2jguJe+obRVvc4TFslseFL03d2TLiDwfPw1Uht1ojJas941jQeZ14D0i/vDOeohGpCdcNw6s976foYT0La59f+UivTmg0U5tfqqGyc/bd1HimhSj6SDw8pf8KroEJ8q1u2WInXzwEcisgRkdrgNebMWo2z790ptQN70p8Xn6nVELALnDqVE3QocTUdWb75UqeWEla+OtUfeI0uV0tolJRxOghav/gQZCcniPCRyvsK5NY7M2ik/cw6O/v75fWFtOQTOe+ai9NPN1Da7UssJuRetFM+rt6/KaH1r2luE5gFrtmp6iE9zS9ZsgLXtLVSftqMD8RWYywfBf9Sq1KJdUsKpUtBdLPYznGNzDJtO4fRkskni5xyKfNwLV8G75bvIEXgdNypK1nxOQE99XFBlnN7nweTvHrnGzkHR6x/B++kWYe5lb4lsgLkRn/nUfS50wnjWdID7iVeUluiT1HBiJe+mHcK1qLfvZOPrTt0JQCcxQzlMrrLINyvqz6NVGZfJHnQfgl6f0hJ9ihOcRJRt3qOUjr94U2s5l1M32ge2ocZbAmvCzSg/lKq0Qr/iAIffxM4EZ69SOv4KHOb9CK1fUUaAeo3XN2uLG1G2dZ/SgugUBzjsyzvBMeLBqKd3Q8Vfyc+++l4aNC3RVbTWS4wHuzJb77vE/83QUMUFDs8ejricY2eJ7PGZkOOOmeKNVm9fQ437115ElnxXocKh75T2VIoTHJ49HOW0R9age2kN+EGp6fQqWFqGspRdKHjy38hsx2c5dbk1zkhE7krXTKCqhcv8u0SlT7wfuwrZN0xE6dqUmB4QRgfHbBONiyQkO9RjnLikN7YJn9/zDdAW4ppu7l+fEZBjIb4XzRcv+JZNdblffFN8Z+eo6SJqQ6uf26NmIrP8275w/XkOrM0GwdwkEmVWbW5P9IfvNHSlWZKE9Iso5B41C6WffC1ST7FWVHAqHblw3Dad3pYJsN8wqR67D9lDJlMHNoude/qVgwhQcxq05rD8oT9ypjwBb8o3CBzL0HS3OFRZKf4fhPJdByg83wPf7oMoo70On/Pw8XZ1Fa1ei4yByeKt5rbWZVk3jIXz7tkEOSBut3q/3AX3UyvhGj8PzuQZ1I8psN84jWDMRv7cF+DhDa8zV/ddBj2KCk5D5KdBZDfDwMyNehKkKwQo/o8kHMnTkcuX+WiX7n7yNbifWS1OMwv++TLyZj8nfl9AGzr38++g8OW1KFq+Dt4vdtMmtPYZSSzdy5lS3OFUiS9y8Nvu2bRd3ErJnbkIrjGzaff/ABx3PgjXpMdppi0lOKtQtPI/KP3oC3g3fCO++sfR19mgMwanLv0a3vpYSEo4hiIy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOxDLgSCwDjsQy4EgsA47EMuBILAOOtAL+D0PoHdDmCF5sAAAAAElFTkSuQmCC";

            var nameElaboro = "C. CASTAÑO";
            var nameAprobo = "C. METRIO";
            var nameReviso = "I. VILLALBA";

            var firmaElaboro = "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";
            var firmaAprobo = "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";
            var firmaReviso = "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";

            var matriculaElaboro = "267773 ANT";
            var matriculaAprobo = "357197 ANT";
            var matriculaReviso = "196375 ANT";

            var nameProyecto = "RENOVACIÓN SUBESTACIÓN BANADÍA 230 kV";
            var codigoDocumento = "CO-RBAN-14113-S-01-D1531";

            var tableFirmas = new List<string[]>()
            {
                new string[8]{ "[B64][WIDTH:1.08][HEIGHT:1.12]" + imagenIEB, "~", "~", "~", "[B64][WIDTH:2.07][HEIGHT:0.83]" + imagenCliente, "~", "~", "~" },
                new string[8]{ "Nombres", "~", "Firma", "Matrícula", "Total de Páginas:", "XXXX", "Fecha de Emisión:", "XXXX.XX.XX" },
                new string[8]{ "Elaboró:", nameElaboro, "[B64][WIDTH:0.45][HEIGHT:0.40]" + firmaElaboro, matriculaElaboro, "Nombre del Proyecto", "~", "~", "~" },
                new string[8]{ "|", "|", "|", "|", nameProyecto, "~", "~", "~" },
                new string[8]{ "Revisó:", nameReviso, "[B64][WIDTH:0.45][HEIGHT:0.40]" + firmaReviso, matriculaReviso, "Código del Documento", "~", "~", "~" },
                new string[8]{ "Aprobó:", nameAprobo, "[B64][WIDTH:0.45][HEIGHT:0.40]" + firmaAprobo, matriculaAprobo, codigoDocumento, "~", "~", "~" },
            };
            #endregion

            body.AppendChild(CreateParagraph.NewP(title1 + sty.SetFontSize(40) + sty.SetBold(), ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));
            
            body.AppendChild(CreateParagraph.NewP(title2 + sty.SetFontSize(40) + sty.SetBold(), ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));
            
            body.AppendChild(CreateParagraph.NewP(title3 + sty.SetFontSize(40) + sty.SetBold(), ParagraphTypes.Custom));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));

            body.AppendChild(CreateTable.NewT(tableAprobacion));
            body.AppendChild(CreateParagraph.NewP(sty.SetFontSize(40), ParagraphTypes.Custom));
            body.AppendChild(CreateTable.NewTImg64(tableFirmas, mainpart, (0, 0)));



            var p = CreateSection.NewS();
            var secProps1 = p.Descendants<SectionProperties>().FirstOrDefault();
            mainpart.Document.Body.AppendChild(p);


            var blanckHeader = mainpart.AddNewPart<HeaderPart>();
            var blanckHeaderPartId = mainpart.GetIdOfPart(blanckHeader);
            new Header().Save(blanckHeader);


            secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = blanckHeaderPartId });
            secProps1.AppendChild(new VerticalTextAlignmentOnPage() { Val = VerticalJustificationValues.Center });

            SizeController.SetPageSize(secProps1, PageSizeTypes.A4, PageOrientationValues.Portrait);
            SizeController.SetMarginSize(secProps1, margin, PageOrientationValues.Portrait);
        }
    }
}
