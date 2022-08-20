using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PruebaControlWord
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

            
            var doc = WordControl.ReadDocument(@"C:\Users\Asus\OneDrive\Desktop\PruebasExcel\test.docx");
            //var doc = new XWPFDocument();
            var datos1 = DatosPrueba();
            var datos2 = DatosPruebav2();
            var datos3 = DatosPruebav3();

            Tuple<int, int> vertical = new Tuple<int, int>(11906, 16838);
            Tuple<int, int> horizontal = new Tuple<int, int>(16838, 11906);

            var seccion1= new CT_SectPr();
            seccion1.pgSz.w = (ulong)vertical.Item1;
            seccion1.pgSz.h = (ulong)vertical.Item2;

            var seccion2 = new CT_SectPr();
            seccion2.pgSz.w = (ulong)horizontal.Item1;
            seccion2.pgSz.h = (ulong)horizontal.Item2;

            // Establece las dimensiones de la primera seccion
            doc.Document.body.sectPr = seccion1;

            WordControl.CreateParagraph(doc, "Nicolas Aguilera Garcia");
            WordControl.CreateParagraph(doc, "Prueba de escritura de parrafos.");
            
            WordControl.CreateParagraph(doc, "Tabla 1.");
            WordControl.CreateTable(doc, datos1, true, true);
            
            WordControl.CreateParagraph(doc, "Tabla 2.", seccion: seccion1);
            WordControl.CreateTable(doc, datos2, false, false);

            // Finaliza la primera seccion
            doc.CreateParagraph().GetCTP().AddNewPPr().sectPr = seccion1;

            // Establece la segunda seccion
            doc.Document.body.sectPr = seccion2;

            WordControl.CreateParagraph(doc, "Tabla 3.");
            WordControl.CreateTable(doc, datos3, true, false);

            WordControl.WriteDocument(@"C:\Users\Asus\OneDrive\Desktop\PruebasExcel\Prueba2.docx", doc);

            
            Console.WriteLine("Finalizado");
            System.Console.WriteLine("Tiempo: " + (DateTime.Now - originalTime));
        }
        public List<string[]> DatosPrueba()
        {
            List<string[]> Datos = new List<string[]>();

            string[] Titulo = new string[1];
            Titulo[0] = "Cara Superior";
            Datos.Add(Titulo);
            string[] Tex1 = new string[4];
            Tex1[0] = "Diseño Losa¬";
            Tex1[1] = "";
            Tex1[2] = "Dir x";
            Tex1[3] = "Dir y";
            Datos.Add(Tex1);

            string[] Tex2 = new string[4];
            Tex2[0] = "Mu negativo¬";
            Tex2[1] = "kg-m/m";
            Tex2[2] = "16541";
            Tex2[3] = "75843";
            Datos.Add(Tex2);

            string[] Tex3 = new string[4];
            Tex3[0] = "Cuantia negativa¬";
            Tex3[1] = "";
            Tex3[2] = "0.0015";
            Tex3[3] = "0.0017";
            Datos.Add(Tex3);

            string[] Tex4 = new string[4];
            Tex4[0] = "Cuantia negativa¬";
            Tex4[1] = "";
            Tex4[2] = "0.0018";
            Tex4[3] = "~";
            Datos.Add(Tex4);

            string[] Tex5 = new string[4];
            Tex5[0] = "Barras¬";
            Tex5[1] = "fi";
            Tex5[2] = "No. 4";
            Tex5[3] = "No. 4";
            Datos.Add(Tex5);

            return Datos;
        }

        public List<string[]> DatosPruebav2()
        {
            List<string[]> Datos = new List<string[]>();

            string[] Titulo = new string[1];
            Titulo[0] = "Diseño Losa";
            Datos.Add(Titulo);

            string[] Tex1 = new string[7];
            Tex1[0] = "";
            Tex1[1] = "";
            Tex1[2] = "";
            Tex1[3] = "";
            Tex1[4] = "";
            Tex1[5] = "";
            Tex1[6] = "";
            Datos.Add(Tex1);

            string[] Titulo2 = new string[1];
            Titulo2[0] = "Datos de diseño";
            Datos.Add(Titulo2);
            string[] Tex11 = new string[7];
            Tex11[0] = "";
            Tex11[1] = "";
            Tex11[2] = "";
            Tex11[3] = "";
            Tex11[4] = "";
            Tex11[5] = "";
            Tex11[6] = "";
            Datos.Add(Tex11);

            string[] Tex2 = new string[7];
            Tex2[0] = "Diseño de losa¬";
            Tex2[1] = "";
            Tex2[2] = "Dir x";
            Tex2[3] = "Dir y";
            Tex2[4] = "";
            Tex2[5] = "";
            Tex2[6] = "";
            Datos.Add(Tex2);

            string[] Tex3 = new string[7];
            Tex3[0] = "Lado de la losa¬";
            Tex3[1] = "mm";
            Tex3[2] = "2300";
            Tex3[3] = "2300";
            Tex3[4] = "Altura efectiva d¬";
            Tex3[5] = "mm";
            Tex3[6] = "175";

            Datos.Add(Tex3);

            string[] Tex4 = new string[7];
            Tex4[0] = "Altura de losa h¬";
            Tex4[1] = "mm";
            Tex4[2] = "250";
            Tex4[3] = "~";
            Tex4[4] = "f'c¬";
            Tex4[5] = "MPa";
            Tex4[6] = "21";
            Datos.Add(Tex4);

            string[] Tex5 = new string[7];
            Tex5[0] = "Recubrimiento r¬";
            Tex5[1] = "mm";
            Tex5[2] = "75";
            Tex5[3] = "~";
            Tex5[4] = "Fy¬";
            Tex5[5] = "MPa";
            Tex5[6] = "420";
            Datos.Add(Tex5);

            return Datos;
        }
        
        static public List<string[]> DatosPruebav3()
        {
            List<string[]> Datos = new List<string[]>();

            string[] Titulo = new string[1];
            Titulo[0] = "RESISTENCIAS NOMINALES";
            Datos.Add(Titulo);

            string[] Tex1 = new string[4];
            Tex1[0] = "Resistencia a cortante del concreto¬";
            Tex1[1] = "Vc";
            Tex1[2] = "kg";
            Tex1[3] = "33037";
            Datos.Add(Tex1);

            string[] Tex2 = new string[4];
            Tex2[0] = "Resistencia a cortante del acero de refuerzo¬";
            Tex2[1] = "Vs";
            Tex2[2] = "kg";
            Tex2[3] = "30398";
            Datos.Add(Tex2);

            string[] Tex3 = new string[4];
            Tex3[0] = "Resistencia a cortante pedestal¬";
            Tex3[1] = "Ø(Vc+Vs)";
            Tex3[2] = "kg";
            Tex3[3] = "47576";
            Datos.Add(Tex3);

            string[] Tex4 = new string[4];
            Tex4[0] = "Separación máxima de refuerzo de cortante¬";
            Tex4[1] = "Smax";
            Tex4[2] = "mm";
            Tex4[3] = "266";
            Datos.Add(Tex4);

            string[] Tex5 = new string[4];
            Tex5[0] = "Resistencia Máxima a tensión del pedestal¬";
            Tex5[1] = "ØTs";
            Tex5[2] = "kg";
            Tex5[3] = "-107350";
            Datos.Add(Tex5);

            string[] Tex6 = new string[4];
            Tex6[0] = "Resistencia Máxima a Compresión del pedestal¬";
            Tex6[1] = "ØTc";
            Tex6[2] = "kg";
            Tex6[3] = "513554";
            Datos.Add(Tex6);

            return Datos;
        }
    }
}
