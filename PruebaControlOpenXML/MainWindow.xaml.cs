using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PruebaControlOpenXML
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DateTime originalTime = DateTime.Now;
            System.Console.WriteLine("Start time: " + originalTime);
            
            //using (var fileDocument = WordprocessingDocument.Create(@"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\TestOpenXML.docx", WordprocessingDocumentType.Document))
            using (var fileDocument = WordprocessingDocument.Create(@"C:\Users\nicho\Downloads\officeTest\TestOpenXML.docx", WordprocessingDocumentType.Document))
            {
                var mainpart = fileDocument.AddMainDocumentPart();
                var doc = mainpart.Document = new Document();

                var body = doc.AppendChild(new Body());
                var c = new WordCommands();

                // Se crean header y footer globales, para tener sus identifucadores cuando sean requeridos
                #region Crear Header Global
                var golbalHeaderPart = mainpart.AddNewPart<HeaderPart>();
                var globalHeaderPartId = mainpart.GetIdOfPart(golbalHeaderPart);

                var globalHeader = c.CreateNewHeaderForSection("UPME 01 – 2018", "DISEÑO DE ESTRUCTURAS DE PÓRTICOS 500 kV");
                globalHeader.Save(golbalHeaderPart);
                #endregion

                #region Crear Footer Global
                var globalFooterPart = mainpart.AddNewPart<FooterPart>();
                var globalFooterPartId = mainpart.GetIdOfPart(globalFooterPart);

                var globalFooter = c.CreateNewFooterForSection("Archivo: CO-TROC-DSIEB-S-00-D1508(3)");
                globalFooter.Save(globalFooterPart);
                #endregion


                // Contenido seccion inicial del documento
                body.AppendChild(c.CreateNewParagraph("1 OBJETO", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Este documento presentar los criterios generales empleados para el análisis y el diseño estructural de los pórticos correspondientes al proyecto Segundo Transformador 500/230/34,5 kV – 360 MVA en la subestación Ocaña 500/230 kV, definido en el “Plan de Expansión de Referencia Generación – Transmisión 2015-2029”. La subestación está localizada en el municipio de Ocaña, departamento de Norte de Santander.", ParagraphTypes.Normal));
                body.AppendChild(c.CreateNewParagraph("Finalmente se presentan los resultados del análisis, el diseño usando el software SAP 2000 y las verificaciones ante las solicitaciones más críticas generadas por las combinaciones de carga.", ParagraphTypes.Normal));

                
                body.AppendChild(c.CreateNewParagraph("2 CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("El diseño de las estructuras metálicas se realizó con base a las especificaciones de las guías de diseño, distancias eléctricas y cargas de conexión, presentadas en los documentos de referencia [1] y [9] y en lo indicado en la referencia [2], considerando las relaciones de esbeltez y los espesores mínimos de los elementos.", ParagraphTypes.Normal));

                
                body.AppendChild(c.CreateNewParagraph("3 MATERIALES", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Se definen en la Tabla 1, con base en lo indicado en la referencia [3]", ParagraphTypes.Normal));
                
                body.AppendChild(c.CreateNewParagraph("Tabla 1 Materiales de los pórticos", ParagraphTypes.Heading1));
                body.AppendChild(c.CreateNewTable(DatosPruebaV1(), haveBorder: true));

                
                body.AppendChild(c.CreateNewParagraph("4 CARGAS ACTUANTES SOBRE LOS PÓRTICOS", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Este documento presentar los criterios generales empleados para el análisis y el diseño estructural de los pórticos correspondientes al proyecto Segundo Transformador 500/230/34,5 kV – 360 MVA en la subestación Ocaña 500/230 kV, definido en el “Plan de Expansión de Referencia Generación – Transmisión 2015-2029”. La subestación está localizada en el municipio de Ocaña, departamento de Norte de Santander.", ParagraphTypes.Normal));
                body.AppendChild(c.CreateNewParagraph("Finalmente se presentan los resultados del análisis, el diseño usando el software SAP 2000 y las verificaciones ante las solicitaciones más críticas generadas por las combinaciones de carga.", ParagraphTypes.Normal));

                
                body.AppendChild(c.CreateNewParagraph("5 COMBINACIONES DE CARGA", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Este documento presentar los criterios generales empleados para el análisis y el diseño estructural de los pórticos correspondientes al proyecto Segundo Transformador 500/230/34,5 kV – 360 MVA en la subestación Ocaña 500/230 kV, definido en el “Plan de Expansión de Referencia Generación – Transmisión 2015-2029”. La subestación está localizada en el municipio de Ocaña, departamento de Norte de Santander.", ParagraphTypes.Normal));
                body.AppendChild(c.CreateNewTable(DatosPruebaV2(), haveBorder: false));
                body.AppendChild(c.CreateNewParagraph("Combinaciones de cargas de servicio (S##):", ParagraphTypes.Normal));
                body.AppendChild(c.CreateNewTable(DatosPruebaV2(), haveBorder: false));

                
                body.AppendChild(c.CreateNewParagraph("6 CRITERIOS DE DEFLEXIÓN", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Las deflexiones de los pórticos se limitaran a los valores indicados en las referencias [3] y [4], con base en lo establecido en la referencia [6]", ParagraphTypes.Normal));
                
                body.AppendChild(c.CreateNewParagraph("Tabla 2 Deflexiones Admisibles", ParagraphTypes.Heading1));
                body.AppendChild(c.CreateNewTable(DatosPruebaV3(), haveBorder: true));
                body.AppendChild(c.CreateNewParagraph("", ParagraphTypes.Table));
                body.AppendChild(c.CreateNewTable(DatosPruebaV4(), haveBorder: true));

                // Asignando propiedades a la seccion inicial
                #region Crear seccion inicial
                var pSection1 = c.CreateNewSection(); // Crea un nuevo parrafo que inicia una seccion
                var secProps1 = pSection1.Descendants<SectionProperties>().FirstOrDefault(); // Obtiene las propiedades de dicha seccion

                body.AppendChild(pSection1); // Agrega el parrafo a la seccion

                secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = globalHeaderPartId });
                secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps1, PageSizeTypes.A4, PageOrientationValues.Portrait);
                WordUtils.SetMarginSize(secProps1, 1984.248, 1984.248, 1984.248, 1984.248, PageOrientationValues.Portrait);
                #endregion




                // Creando pagina de titulo para anexos
                c.CreateNewSectionDivider(ref mainpart, "ANEXO 2: CÁLCULO ESTRUCTURAL COLUMNAS C7 TORRECILLAS SOBRE MURO CORTAFUEGO.", (1984.248, 1984.248, 1984.248, 1984.248));

                // Contenido de la seccion de anexos
                body.AppendChild(c.CreateNewParagraph("6 CRITERIOS DE DEFLEXIÓN", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewParagraph("Las deflexiones de los pórticos se limitaran a los valores indicados en las referencias [3] y [4], con base en lo establecido en la referencia [6]", ParagraphTypes.Normal));

                body.AppendChild(c.CreateNewParagraph("Tabla 2 Deflexiones Admisibles", ParagraphTypes.Heading1));
                body.AppendChild(c.CreateNewTable(DatosPruebaV3(), haveBorder: true));
                body.AppendChild(c.CreateNewParagraph("", ParagraphTypes.Table));
                body.AppendChild(c.CreateNewTable(DatosPruebaV4(), haveBorder: true));


                #region Crear Header seccion anexos
                var anexoHeaderPart = mainpart.AddNewPart<HeaderPart>();
                var anexoHeaderPartId = mainpart.GetIdOfPart(anexoHeaderPart);

                var anexoHeader = c.CreateNewHeaderForSection("UPME 01 – 2018", "ANEXO 2");
                anexoHeader.Save(anexoHeaderPart);
                #endregion

                // Asignando propiedades a la seccion final (en este caso la de anexos)
                #region Crear final seccion
                var secProps2 = c.CreateFinalSection();
                body.AppendChild(secProps2);

                secProps2.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = anexoHeaderPartId });
                secProps2.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps2, PageSizeTypes.A4, PageOrientationValues.Portrait);
                WordUtils.SetMarginSize(secProps2, 1984.248, 1984.248, 1984.248, 1984.248, PageOrientationValues.Portrait);
                #endregion



                //////
                ///
                /// Nota:
                /// Definir una seccion se realiza despues de haber agregado dentro del -body- el contenido que le corresponde a dicha seccion.
                /// 
                //////

                
                mainpart.Document.Save(); // El guardado es opcional, por defecto OpenXML realiza autoguardado de cada cambio realizado
            }

            System.Console.WriteLine("Final time: " + (DateTime.Now - originalTime));
        }

        #region Generar datos de prueba
        public List<string[]> DatosPruebaV1()
        {
            var datos = new List<string[]>();

            datos.Add(new string[3]
            {
                "Elemento", "Perfiles¬", "ASTM A-572 Gr50¬"
            });
            datos.Add(new string[3]
            {
                "|", "Platinas¬", "ASTM A-36¬"
            });
            datos.Add(new string[3]
            {
                "|", "Soldadura¬", "De acuerdo AWS D1.1 y D1.3.\r\nElectrodos E70-XX¬"
            });
            datos.Add(new string[3]
            {
                "|", "Tornillos¬", "ASTM A-394 TIPO 0¬"
            });
            datos.Add(new string[3]
            {
                "|", "Pernos de anclaje¬", "F1554 Gr 55¬"
            });
            datos.Add(new string[3]
            {
                "|", "Arandelas¬", "ASTM F-436¬"
            });
            datos.Add(new string[3]
            {
                "|", "Tuercas¬", "ASTM A-563¬"
            });
            datos.Add(new string[3]
            {
                "|", "Galvanización¬", "ASTM A-123¬"
            });


            return datos;
        }

        public List<string[]> DatosPruebaV2()
        {
            var datos = new List<string[]>();

            datos.Add(new string[2]
            {
                "U1.0¬", "1,2PP + 1,3CT + 1,0CMM¬"
            });
            datos.Add(new string[2]
            {
                "U2.0¬", "1,1PP +1,38Vx + 1,38CTv + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U2.1¬", "1,1PP + 1,38Vy + 1,38CTv + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U4.0¬", "1,1PP + 1,0CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U5.0¬", "1,1PP + 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U5.1¬", "1,1PP - 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U1.0¬", "1,2PP + 1,3CT + 1,0CMM¬"
            });
            datos.Add(new string[2]
            {
                "U2.0¬", "1,1PP +1,38Vx + 1,38CTv + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U2.1¬", "1,1PP + 1,38Vy + 1,38CTv + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U4.0¬", "1,1PP + 1,0CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U5.0¬", "1,1PP + 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT¬"
            });
            datos.Add(new string[2]
            {
                "U5.1¬", "1,1PP - 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT¬"
            });


            return datos;
        }

        public List<string[]> DatosPruebaV3()
        {
            var datos = new List<string[]>();

            datos.Add(new string[5]
            {
                "TIPO DE DEFLEXIÓN[N]", "ESTRUCTURA CLASE A[N]", "~", "ESTRUCTURA CLASE B[N]", "~"
            });
            datos.Add(new string[5]
            {
                "|", "Elementos Horizontales", "Elementos Verticales", "Elementos Horizontales", "Elementos Verticales"
            });
            datos.Add(new string[5]
            {
                "Horizontal", "1/200", "1/100", "1/100", "1/100"
            });
            datos.Add(new string[5]
            {
                "Vertical", "1/200", "", "1/200", ""
            });

            return datos;
        }

        public List<string[]> DatosPruebaV4()
        {
            var datos = new List<string[]>();

            datos.Add(new string[2]
            {
                "Clasificación de los miembros, según ASCE - 113[N]", "~"
            });
            datos.Add(new string[2]
            {
                "Clase A", "Cuando existan equipos sobre los pórticos."
            });
            datos.Add(new string[2]
            {
                "Clase B", "Cuando no existen equipos sobre los pórticos."
            });

            return datos;
        }
        #endregion
    }
}
