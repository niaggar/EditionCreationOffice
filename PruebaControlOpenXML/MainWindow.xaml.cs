using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
using Microsoft.Win32;

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

        public string GetSaveRoute()
        {
            var createFile = new SaveFileDialog()
            {
                FileName = "TestOpenXML.docx",
                Filter = "Word Files (*.docx)|*.docx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                DefaultExt = "docx"
            };
            var res = createFile.ShowDialog();
            if (res != true) return "";

            return createFile.FileName;
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            string route = GetSaveRoute();
            if (route == "") return;


            DateTime originalTime = DateTime.Now;
            System.Console.WriteLine("Start time: " + originalTime);


            // 1 Pulgada = 1440 twips
            // 1 cm = 567 twips


            var documentSize = PageSizeTypes.A4;
            var documentMargins = (top: 1, right: 1, bottom: 1, left: 1);
            var c = new WordCommands();
            
            var fileDocument = c.CreateDocument(route);
            if (fileDocument == null)
            {
                System.Console.WriteLine("End time: " + DateTime.Now);
                System.Console.WriteLine("Elapsed time: " + (DateTime.Now - originalTime));
                return;
            }

            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());


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

            var name = System.IO.Path.GetFileNameWithoutExtension(route);
            var globalFooter = c.CreateNewFooterForSection($"Archivo: {name}");
            globalFooter.Save(globalFooterPart);
            #endregion


            var ciclos = 1;
            for (int i = 0; i < ciclos; i++)
            {
                // Asignando propiedades a la seccion inicial
                #region Crear seccion inicial
                var pSection1 = c.CreateNewSection(); // Crea un nuevo parrafo que inicia una seccion
                var secProps1 = pSection1.Descendants<SectionProperties>().FirstOrDefault(); // Obtiene las propiedades de dicha seccion

                secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = globalHeaderPartId });
                secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps1, documentSize, PageOrientationValues.Portrait);
                WordUtils.SetMarginSize(secProps1, documentMargins, PageOrientationValues.Portrait);
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
                body.AppendChild(c.CreateNewParagraph("", ParagraphTypes.Table));
                body.AppendChild(c.CreateNewTable(DatosPruebaDobleTitulo(), haveBorder: true));


                // Agregando primera seccion
                body.AppendChild(pSection1);

                #region Crear Header seccion anexos
                var anexoHeaderPart = mainpart.AddNewPart<HeaderPart>();
                var anexoHeaderPartId = mainpart.GetIdOfPart(anexoHeaderPart);

                var anexoHeader = c.CreateNewHeaderForSection("UPME 01 – 2018", "ANEXO 2");
                anexoHeader.Save(anexoHeaderPart);
                #endregion

                // Creando pagina de titulo para anexos
                c.CreateNewSectionDivider(ref mainpart, "ANEXO 2: CÁLCULO ESTRUCTURAL COLUMNAS C7 TORRECILLAS SOBRE MURO CORTAFUEGO.", documentMargins);

                // Asignando propiedades a la segunda
                #region Crear segunda seccion
                var pSection2 = c.CreateNewSection(); // Crea un nuevo parrafo que inicia una seccion
                var secProps2 = pSection2.Descendants<SectionProperties>().FirstOrDefault(); // Obtiene las propiedades de dicha seccion

                secProps2.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = anexoHeaderPartId });
                secProps2.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps2, documentSize, PageOrientationValues.Landscape);
                WordUtils.SetMarginSize(secProps2, documentMargins, PageOrientationValues.Landscape);
                #endregion

                var pagesizeSec2 = WordUtils.GetPaperSize(documentSize);
                var marginsSec2 = documentMargins;
                var widthUtilSpaceSec2 = pagesizeSec2.width - marginsSec2.right - marginsSec2.left;
                var heightUtilSpaceSec2 = pagesizeSec2.height - marginsSec2.top - marginsSec2.bottom;

                // Contenido de la seccion de anexos
                var r = @"C:\Users\Nicolas\Desktop\";
                //var img1 = c.CreateNewImage(mainpart, r + "img.jpg", escale: 0.5);
                var img2 = c.CreateNewImage(mainpart, r + "img.jpg", width: widthUtilSpaceSec2, height: heightUtilSpaceSec2);
                //var img3 = c.CreateNewImage(mainpart, r + "Diagonal.jpeg", height: heightUtilSpaceSec2);

                //body.AppendChild(c.CreateNewParagraph("Silueta", ParagraphTypes.Heading2));
                //body.AppendChild(img1);
                //body.AppendChild(c.CreateNewPargraphPageBreak());

                //body.AppendChild(c.CreateNewParagraph("Vista lateral", ParagraphTypes.Heading2));
                body.AppendChild(img2);
                body.AppendChild(c.CreateNewPargraphPageBreak());

                //body.AppendChild(c.CreateNewParagraph("Vista Frontal", ParagraphTypes.Heading2));
                //body.AppendChild(img3);

                // Agregando primera seccion
                body.AppendChild(pSection2);

                #region Crear tercera seccion
                var pSection3 = c.CreateNewSection(); // Crea un nuevo parrafo que inicia una seccion
                var secProps3 = pSection3.Descendants<SectionProperties>().FirstOrDefault(); // Obtiene las propiedades de dicha seccion

                secProps3.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = anexoHeaderPartId });
                secProps3.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps3, documentSize, PageOrientationValues.Portrait);
                WordUtils.SetMarginSize(secProps3, documentMargins, PageOrientationValues.Portrait);
                #endregion

                var pagesizeSecFinal = WordUtils.GetPageSize(secProps3);
                var marginsSecFinal = WordUtils.GetMarginSize(secProps3);
                var widthUtilSpaceSecFinal = (long)WordUtils.ConvertTwipToCm(pagesizeSecFinal.width - marginsSecFinal.left - marginsSecFinal.right);
                var heightUtilSpaceSecFinal = (long)WordUtils.ConvertTwipToCm(pagesizeSecFinal.height - marginsSecFinal.top - marginsSecFinal.bottom);

                // Contenido de tercera seccion
                body.AppendChild(c.CreateNewParagraph("Identificación de nodos y elementos del soporte", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewImageTable(DatosPruebaV5(), (widthUtilSpaceSecFinal, heightUtilSpaceSecFinal), mainpart));

                body.AppendChild(pSection3);
            }



            // Asignando propiedades a la seccion final
            #region Crear final seccion
            var secFinal = c.CreateFinalSection();
            
            WordUtils.SetPageSize(secFinal, documentSize, PageOrientationValues.Portrait);
            WordUtils.SetMarginSize(secFinal, documentMargins, PageOrientationValues.Portrait);
            #endregion

            // Agregando seccion final
            body.AppendChild(secFinal);




            //////
            ///
            /// Nota:
            /// Definir una seccion se realiza despues de haber agregado dentro del -body- el contenido que le corresponde a dicha seccion.
            /// 
            //////



            mainpart.Document.Save(); // El guardado es opcional, por defecto OpenXML realiza autoguardado de cada cambio realizado
            fileDocument.Close();


            //WordUtils.SaveDocumentAsPdf(route, System.IO.Path.ChangeExtension(route, ".pdf"));


            System.Console.WriteLine("Final time: " + (DateTime.Now - originalTime));
        }

        #region Generar datos de prueba
        public List<string[]> DatosPruebaV1()
        {
            var datos = new List<string[]>();

            datos.Add(new string[3]
            {
                "Elemento", $"Perfiles{WordUtils.SetLeftAligment()}", $"ASTM A-572 Gr50{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Platinas{WordUtils.SetLeftAligment()}", $"ASTM A-36{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Soldadura{WordUtils.SetLeftAligment()}", $"De acuerdo AWS D1.1 y D1.3.\r\nElectrodos E70-XX{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Tornillos{WordUtils.SetLeftAligment()}", $"ASTM A-394 TIPO 0{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Pernos de anclaje{WordUtils.SetLeftAligment()}", $"F1554 Gr 55{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Arandelas{WordUtils.SetLeftAligment()}", $"ASTM F-436{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Tuercas{WordUtils.SetLeftAligment()}", $"ASTM A-563{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[3]
            {
                "|", $"Galvanización{WordUtils.SetLeftAligment()}", $"ASTM A-123{WordUtils.SetLeftAligment()}"
            });


            return datos;
        }

        public List<string[]> DatosPruebaV2()
        {
            var datos = new List<string[]>();

            datos.Add(new string[2]
            {
                $"U1.0{WordUtils.SetLeftAligment()}", $"1,2PP + 1,3CT + 1,0CMM{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U2.0{WordUtils.SetLeftAligment()}", $"1,1PP +1,38Vx + 1,38CTv + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U2.1{WordUtils.SetLeftAligment()}", $"1,1PP + 1,38Vy + 1,38CTv + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U4.0{WordUtils.SetLeftAligment()}", $"1,1PP + 1,0CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U5.0{WordUtils.SetLeftAligment()}", $"1,1PP + 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U5.1{WordUtils.SetLeftAligment()}", $"1,1PP - 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U1.0{WordUtils.SetLeftAligment()}", $"1,2PP + 1,3CT + 1,0CMM{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U2.0{WordUtils.SetLeftAligment()}", $"1,1PP +1,38Vx + 1,38CTv + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U2.1{WordUtils.SetLeftAligment()}", $"1,1PP + 1,38Vy + 1,38CTv + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U4.0{WordUtils.SetLeftAligment()}", $"1,1PP + 1,0CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U5.0{WordUtils.SetLeftAligment()}", $"1,1PP + 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });
            datos.Add(new string[2]
            {
                $"U5.1{WordUtils.SetLeftAligment()}", $"1,1PP - 1,0Ex + 0,3Ey - 1,0Ez + 0,75CC + 1,1CT{WordUtils.SetLeftAligment()}"
            });


            return datos;
        }

        public List<string[]> DatosPruebaV3()
        {
            var datos = new List<string[]>();

            datos.Add(new string[5]
            {
                $"TIPO DE DEFLEXIÓN{WordUtils.SetBold()}", $"ESTRUCTURA CLASE A{WordUtils.SetBold()}", "~", $"ESTRUCTURA CLASE B{WordUtils.SetBold()}", "~"
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
                $"Clasificación de los miembros, según ASCE - 113{WordUtils.SetBold()}", "~"
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

        public List<string[]> DatosPruebaV5()
        {
            var datos = new List<string[]>();

            datos.Add(new string[2]
            {
                $"COLUMNA C3{WordUtils.SetBold()}", "~"
            });
            datos.Add(new string[2]
            {
                @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\img2\tabla1.jpeg", @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\img2\tabla2.jpeg"
            });

            return datos;
        }

        public List<string[]> DatosPruebaDobleTitulo()
        {
            List<string[]> Datos = new List<string[]>();

            var stylesTitle = $"{WordUtils.SetBold()}{WordUtils.SetFontSize(24)}{WordUtils.SetCellColor("#FF0000")}{WordUtils.SetFontColor("#FFFFFF")}";

            string[] Titulo = new string[4];
            Titulo[0] = $"Cara Superior{stylesTitle}";
            Titulo[1] = "~";
            Titulo[2] = "~";
            Titulo[3] = "~";
            Datos.Add(Titulo);

            string[] Tex1 = new string[4];
            Tex1[0] = $"Diseño Losa{WordUtils.SetLeftAligment()}{stylesTitle}";
            Tex1[1] = $"{stylesTitle}";
            Tex1[2] = $"Dir x{stylesTitle}";
            Tex1[3] = $"Dir y{stylesTitle}";
            Datos.Add(Tex1);

            string[] Tex2 = new string[4];
            Tex2[0] = $"Mu negativo{WordUtils.SetLeftAligment()}";
            Tex2[1] = "kg-m/m";
            Tex2[2] = "16541";
            Tex2[3] = "75843";
            Datos.Add(Tex2);

            string[] Tex3 = new string[4];
            Tex3[0] = $"Cuantia negativa{WordUtils.SetLeftAligment()}";
            Tex3[1] = "";
            Tex3[2] = "0.0015";
            Tex3[3] = "0.0017";
            Datos.Add(Tex3);

            string[] Tex4 = new string[4];
            Tex4[0] = $"Cuantia negativa{WordUtils.SetLeftAligment()}";
            Tex4[1] = "";
            Tex4[2] = "0.0018";
            Tex4[3] = "~";
            Datos.Add(Tex4);

            string[] Tex5 = new string[4];
            Tex5[0] = $"Barras{WordUtils.SetLeftAligment()}";
            Tex5[1] = "fi";
            Tex5[2] = "No. 4";
            Tex5[3] = "No. 4";
            Datos.Add(Tex5);

            return Datos;
        }
        #endregion



        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var count = 2000;


            DateTime firstTime = DateTime.Now;
            Console.WriteLine("Start time: " + firstTime);

            var r = @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\img2\";
            ImageControl.CopyImage(r + "Conexion2.jpeg", count);

            Console.WriteLine("Final time copy file: " + (DateTime.Now - firstTime));



            DateTime secondTime = DateTime.Now;
            Console.WriteLine("Start time: " + secondTime);

            var img = new BitmapImage();
            var imgBase64 = "";
            using (var fs = new FileStream(r + "Conexion2.jpeg", FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                img.BeginInit();
                img.StreamSource = fs;
                imgBase64 = ImageControl.ImageToBase64(img);
                img.EndInit();
            }

            string[] datos = new string[count];
            for (int i = 0; i < count; i++)
            {
                datos[i] = "[jpeg]" + imgBase64;
            }

            ImageControl.WriteTextToFile(datos);

            Console.WriteLine("Final time convert to txt: " + (DateTime.Now - secondTime));
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string route = GetSaveRoute();
            if (route == "") return;


            
            DateTime firstTime = DateTime.Now;
            Console.WriteLine("Start time: " + firstTime);
            
            

            var c = new WordCommands();
            var fileDocument = c.CreateDocument(route);
            if (fileDocument == null)
            {
                return;
            }

            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());

            string[] readText = File.ReadAllLines(@"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\images\img.txt");

            foreach (var item in readText)
            {
                body.AppendChild(c.CreateNewParagraph("Silueta", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewBase64Image(mainpart, item, escale: 0.5));
                body.AppendChild(c.CreateNewPargraphPageBreak());
            }

            mainpart.Document.Save();
            fileDocument.Close();



            Console.WriteLine("Final time base64: " + (DateTime.Now - firstTime));
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            string route = GetSaveRoute();
            if (route == "") return;



            DateTime firstTime = DateTime.Now;
            Console.WriteLine("Start time: " + firstTime);



            var c = new WordCommands();
            var fileDocument = c.CreateDocument(route);
            if (fileDocument == null)
            {
                return;
            }

            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());

            var baseRoute = @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\images\";
            for (int i = 0; i < 2000; i++)
            {
                body.AppendChild(c.CreateNewParagraph("Silueta", ParagraphTypes.Heading2));
                body.AppendChild(c.CreateNewImage(mainpart, baseRoute + $"img-{i}.jpeg", escale: 0.5));
                body.AppendChild(c.CreateNewPargraphPageBreak());
            }

            mainpart.Document.Save();
            fileDocument.Close();



            Console.WriteLine("Final time files: " + (DateTime.Now - firstTime));
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            string route = GetSaveRoute();
            if (route == "") return;

            var documentSize = PageSizeTypes.A4;
            var documentMargins = (top: 0.79, right: 0.98, bottom: 1.18, left: 1.38);

            var c = new WordCommands();

            var fileDocument = c.CreateDocument(route);
            var mainpart = fileDocument.AddMainDocumentPart();
            var doc = mainpart.Document = new Document();
            var body = doc.AppendChild(new Body());



            #region Crear Header Global
            var golbalHeaderPart = mainpart.AddNewPart<HeaderPart>();
            var globalHeaderPartId = mainpart.GetIdOfPart(golbalHeaderPart);

            var globalHeader = c.CreateNewHeaderForSection("CO-RBAN: RENOVACIÓN SUBESTACIÓN BANADÍA 230 kV", "MEMORIA DE DISEÑO DE ESTRUCTURAS METÁLICAS DE PÓRTICOS");
            globalHeader.Save(golbalHeaderPart);
            #endregion

            #region Crear Footer Global
            var globalFooterPart = mainpart.AddNewPart<FooterPart>();
            var globalFooterPartId = mainpart.GetIdOfPart(globalFooterPart);

            var name = System.IO.Path.GetFileNameWithoutExtension(route);
            var globalFooter = c.CreateNewFooterForSection($"Archivo: {name}");
            globalFooter.Save(globalFooterPart);
            #endregion




            #region Crear Tabla de contenidos
            var sdtBlock = new SdtBlock();
            sdtBlock.InnerXml = c.GetTOC("Contenido", 16);
            doc.MainDocumentPart.Document.Body.AppendChild(sdtBlock);


            var settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings { BordersDoNotSurroundFooter = new BordersDoNotSurroundFooter() { Val = true } };
            settingsPart.Settings.Append(new UpdateFieldsOnOpen() { Val = true });
            #endregion

            #region Estilos
            StyleDefinitionsPart part = doc.MainDocumentPart.StyleDefinitionsPart;
            if (part == null) part = StyleGenerator.AddStylesPartToPackage(doc);
            StyleGenerator.CreateAndAddParagraphStyle(part);
            #endregion

            #region Create Numberingo
            var numbering = mainpart.AddNewPart<NumberingDefinitionsPart>();

            Numbering globalNumbering = new Numbering();
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




            var abs = new AbstractNum() { AbstractNumberId = 0 };


            Nsid nsid3 = new Nsid() { Val = "70913756" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "624EA66A" };
            AbstractNumDefinitionName abstractNumDefinitionName1 = new AbstractNumDefinitionName() { Val = "TitlesNumberingDEFAULT" };


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
            level1.Append(new ParagraphStyleIdInLevel() { Val = "tt1" });
            level1.Append(pPr1);
            level1.Append(rPr1);


            var level2 = new Level() { LevelIndex = 1 };

            var pPr2 = new PreviousParagraphProperties();
            pPr2.Append(new Tabs(new TabStop() { Val = TabStopValues.Number, Position = 0 }));
            pPr2.Append(new Indentation() { Left = "576", Hanging = "576" });

            var rPr2 = new NumberingSymbolRunProperties();
            rPr2.Append(new RunFonts() { Hint = FontTypeHintValues.Default });

            level2.Append(new NumberingFormat() { Val = NumberFormatValues.Decimal });
            level2.Append(new ParagraphStyleIdInLevel() { Val = "tt2" });
            level2.Append(new StartNumberingValue() { Val = 1 });
            level2.Append(new LevelText() { Val = "%1.%2" });
            level2.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
            level2.Append(pPr2);
            level2.Append(rPr2);


            abs.Append(nsid3);
            abs.Append(multiLevelType3);
            abs.Append(templateCode3);
            abs.Append(abstractNumDefinitionName1);
            abs.Append(level1);
            abs.Append(level2);

            for (int i = 2; i < 9; i++)
            {
                var level = new Level() { LevelIndex = i };

                var pPr = new PreviousParagraphProperties();
                pPr.Append(new Tabs(new TabStop() { Val = TabStopValues.Number, Position = 0 }));
                pPr.Append(new Indentation() { Left = $"{577 + (144 * (i - 1))}", Hanging = $"{577 + (144 * (i - 1))}" });

                var rPr = new NumberingSymbolRunProperties();
                rPr.Append(new RunFonts() { Hint = FontTypeHintValues.Default });

                level.Append(new NumberingFormat() { Val = NumberFormatValues.Decimal });
                level.Append(new ParagraphStyleIdInLevel() { Val = $"tt{i+1}" });
                level.Append(new StartNumberingValue() { Val = 1 });
                level.Append(new LevelText() { Val = $"%{i}" });
                level.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
                level.Append(pPr);
                level.Append(rPr);

                abs.Append(level);
            }


            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = 1 };
            numberingInstance.Append(new AbstractNumId() { Val = 0 });
            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 15 };
            numberingInstance.Append(new AbstractNumId() { Val = 0 });

            globalNumbering.Append(abs);
            globalNumbering.Append(numberingInstance);
            globalNumbering.Append(numberingInstance2);
            numbering.Numbering = globalNumbering;
            #endregion





            #region Seccion 1
            #region Crear seccion inicial
            // Crear seccion
            var pSection1 = c.CreateNewSection();
            var secProps1 = pSection1.Descendants<SectionProperties>().FirstOrDefault();

            // Agregar header y footer de seccion
            secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = globalHeaderPartId });
            secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });

            // Establecer tamaños
            WordUtils.SetPageSize(secProps1, documentSize, PageOrientationValues.Portrait);
            WordUtils.SetMarginSize(secProps1, documentMargins, PageOrientationValues.Portrait);
            #endregion

            #region Contenido
            body.AppendChild(c.CreateNewParagraph("OBJETO", ParagraphTypes.Heading1));
            body.AppendChild(c.CreateNewParagraph("Este documento presentar los criterios generales empleados para el análisis y el diseño estructural de los pórticos correspondientes al proyecto Segundo Transformador 500/230/34,5 kV – 360 MVA en la subestación Ocaña 500/230 kV, definido en el “Plan de Expansión de Referencia Generación – Transmisión 2015-2029”. La subestación está localizada en el municipio de Ocaña, departamento de Norte de Santander.", ParagraphTypes.Normal));
            body.AppendChild(c.CreateNewParagraph("OBJETO", ParagraphTypes.Heading1));
            body.AppendChild(c.CreateNewParagraph("Finalmente se presentan los resultados del análisis, el diseño usando el software SAP 2000 y las verificaciones ante las solicitaciones más críticas generadas por las combinaciones de carga.", ParagraphTypes.Normal));
            body.AppendChild(c.CreateNewParagraph("CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2));
            body.AppendChild(c.CreateNewParagraph("El diseño de las estructuras metálicas se realizó con base a las especificaciones de las guías de diseño, distancias eléctricas y cargas de conexión, presentadas en los documentos de referencia [1] y [9] y en lo indicado en la referencia [2], considerando las relaciones de esbeltez y los espesores mínimos de los elementos.", ParagraphTypes.Normal));
            body.AppendChild(c.CreateNewParagraph("OBJETO", ParagraphTypes.Heading1));
            body.AppendChild(c.CreateNewParagraph("OBJETO", ParagraphTypes.Heading1));
            body.AppendChild(c.CreateNewParagraph("CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2));
            body.AppendChild(c.CreateNewParagraph("CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2));
            body.AppendChild(c.CreateNewParagraph("CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2));
            body.AppendChild(c.CreateNewParagraph("OBJETO", ParagraphTypes.Heading1));


            body.AppendChild(c.CreateNewParagraph("Tabla 1 Materiales de los pórticos", ParagraphTypes.Heading1));


            var p = new Paragraph();

            var run = new Run();
            var runStyle = new StyleRunProperties();
            var paragraphStyle = new ParagraphProperties();

            paragraphStyle.ParagraphStyleId = new ParagraphStyleId() { Val = "Caption" };
            abs.AppendChild(p);




            //p.InnerXml = $@"
            //<w:p>
            //    <w:pPr>
            //        <w:pStyle w:val=""caption"" />
            //        <w:keepNext />
            //    </w:pPr>
            //    <w:r>
            //        <w:t xml:space=""preserve"">Table </w:t>
            //    </w:r>
            //    <w:fldSimple w:instr="" SEQ Table \* ARABIC "">
            //    </w:fldSimple>
            //    <w:r>
            //        <w:t xml:space=""preserve"">. </w:t>
            //    </w:r>
            //    <w:proofErr w:type=""spellStart"" />
            //    <w:r>
            //        <w:t>abcdefghijk</w:t>
            //    </w:r>
            //    <w:proofErr w:type=""spellEnd"" />
            //</w:p>";
            body.AppendChild(p);

            body.AppendChild(c.CreateNewTable(DatosPruebaV1(), haveBorder: true));
            #endregion

            body.AppendChild(pSection1);
            #endregion


            #region Seccion final
            #region Crear final seccion
            var secFinal = c.CreateFinalSection();

            WordUtils.SetPageSize(secFinal, documentSize, PageOrientationValues.Portrait);
            WordUtils.SetMarginSize(secFinal, documentMargins, PageOrientationValues.Portrait);
            #endregion

            // Agregando seccion final
            body.AppendChild(secFinal);
            #endregion


            mainpart.Document.Save();
            fileDocument.Close();
        }
    }
}
