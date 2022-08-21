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

                #region nameSpaces
                doc.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                doc.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
                doc.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
                doc.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
                doc.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
                doc.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
                doc.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
                doc.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
                doc.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
                doc.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                doc.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
                doc.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
                doc.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                doc.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                doc.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                doc.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                doc.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                doc.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                doc.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                doc.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                doc.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                doc.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                doc.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                doc.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                doc.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                doc.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                doc.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                doc.AddNamespaceDeclaration("Ignorable", "w14 w15 w16se w16cid wp14");
                doc.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                #endregion



                var body = doc.AppendChild(new Body());
                var control = new CreateTable();

                
                
                #region Crear Header Global
                var golbalHeaderPart = mainpart.AddNewPart<HeaderPart>();
                var globalHeaderPartId = mainpart.GetIdOfPart(golbalHeaderPart);

                var globalHeader = control.CreateHeaderForSection("UPME 01 – 2018", "DISEÑO DE ESTRUCTURAS DE PÓRTICOS 500 kV");
                globalHeader.Save(golbalHeaderPart);
                #endregion

                #region Crear Footer Global
                var globalFooterPart = mainpart.AddNewPart<FooterPart>();
                var globalFooterPartId = mainpart.GetIdOfPart(globalFooterPart);

                var globalFooter = control.CreateFooterForSection("Archivo: CO-TROC-DSIEB-S-00-D1508(3)");
                globalFooter.Save(globalFooterPart);
                #endregion




                var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec at mauris suscipit, bibendum justo vel, tempus quam. Nam vitae faucibus sem. Proin a odio a sapien blandit tristique a a enim. Sed elementum lectus sed est facilisis, a placerat erat consectetur. Morbi vitae molestie elit, eget congue lacus. Ut vitae pellentesque ligula, pellentesque maximus neque. Etiam enim metus, tristique non est sed, finibus venenatis purus. Aliquam maximus leo nec maximus cursus. Suspendisse massa arcu, efficitur feugiat sapien sed, imperdiet laoreet odio. Praesent vehicula vehicula viverra. Proin sollicitudin tellus non sem scelerisque, quis eleifend est rutrum. Interdum et malesuada fames ac ante ipsum primis in faucibus. Proin efficitur consequat nisi, ornare consectetur nisi placerat eu. Suspendisse potenti. Suspendisse posuere hendrerit finibus. Suspendisse condimentum tellus non dapibus consequat.";

                var head1 = control.CrearNuevoParrafo("LISTA DE ANEXOS", ParagraphTypes.Heading1);
                var parr1 = control.CrearNuevoParrafo(text, ParagraphTypes.Normal);
                var head2 = control.CrearNuevoParrafo("2 CRITERIOS Y ANÁLISIS DE DISEÑO", ParagraphTypes.Heading2);
                var parr2 = control.CrearNuevoParrafo(text, ParagraphTypes.Normal);
                var table = control.CrearNuevaTablaWord();
                body.AppendChild(head1);
                body.AppendChild(parr1);
                body.AppendChild(table);
                body.AppendChild(head2);
                body.AppendChild(parr2);
                

                

                #region Crear nueva seccion
                var pSection1 = control.CreateNewSection(); // Crea un nuevo parrafo que inicia una seccion
                var secProps1 = pSection1.Descendants<SectionProperties>().FirstOrDefault(); // Obtiene las propiedades de dicha seccion
                
                body.AppendChild(pSection1); // Agrega el parrafo a la seccion

                secProps1.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = globalHeaderPartId });
                secProps1.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps1, PageSizeTypes.A4, PageOrientationValues.Landscape);
                WordUtils.SetMarginSize(secProps1, 1984.248, 1984.248, 1984.248, 1984.248, PageOrientationValues.Landscape);
                #endregion



                var head3 = control.CrearNuevoParrafo("3 CARGAS ACTUANTES SOBRE LOS PÓRTICOS", ParagraphTypes.Heading2);
                var parr3 = control.CrearNuevoParrafo(text, ParagraphTypes.Normal);
                var head4 = control.CrearNuevoParrafo("4 COMBINACIONES DE CARGA", ParagraphTypes.Heading2);
                var parr4 = control.CrearNuevoParrafo(text, ParagraphTypes.Normal);
                body.AppendChild(head3);
                body.AppendChild(parr3);
                body.AppendChild(head4);
                body.AppendChild(parr4);
                


                
                #region Crear final seccion
                var secProps2 = control.CreateFinalSection(); // Crea un nuevo parrafo que inicia una seccion

                body.AppendChild(secProps2); // Agrega el parrafo a la seccion

                secProps2.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = globalHeaderPartId });
                secProps2.AppendChild(new FooterReference() { Type = HeaderFooterValues.Default, Id = globalFooterPartId });
                WordUtils.SetPageSize(secProps2, PageSizeTypes.A4, PageOrientationValues.Portrait);
                WordUtils.SetMarginSize(secProps2, 1984.248, 1984.248, 1984.248, 1984.248, PageOrientationValues.Portrait);
                #endregion






                mainpart.Document.Save();
            }

            System.Console.WriteLine("Final time: " + (DateTime.Now - originalTime));
        }
    }
}
