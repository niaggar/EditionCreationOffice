using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace PruebaControlExcel
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

        static void writeToExcel()
        {
            Excel.Application myexcelApplication = new Excel.Application();
            if (myexcelApplication != null)
            {
                Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();
                DateTime originalTime = DateTime.Now;
                System.Console.WriteLine("Start time interop: " + originalTime);
                int i = 0;
                int rowLimit = 65535;
                string[,] arr = new string[rowLimit, 10];
                for (i = 0; i < rowLimit; i++)
                {
                    arr[i, 0] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 1] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 2] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 3] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 4] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 5] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 6] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 7] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 8] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                    arr[i, 9] = "ZOMG PLEASE SURVIVE THIS STRESS TEST";
                }
                Excel.Range c1 = (Excel.Range)myexcelWorksheet.Cells[1, 1];
                Excel.Range c2 = (Excel.Range)myexcelWorksheet.Cells[rowLimit, 10];//fila columna
                Excel.Range range = myexcelWorksheet.get_Range(c1, c2);
                range.Value2 = arr;
                myexcelApplication.ActiveWorkbook.SaveAs(@"C:\Users\Asus\OneDrive\Desktop\PruebasExcel\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                myexcelWorkbook.Close();
                myexcelApplication.Quit();
                Console.WriteLine("Finalizado con Interop");
                System.Console.WriteLine("tiempo tomado Interop: " + (DateTime.Now - originalTime));
            }
        }

        static void Test()
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("StressTest");
            int i = 0;
            int rowLimit = 65535;
            DateTime originalTime = DateTime.Now;
            System.Console.WriteLine("Start time libreria: " + originalTime);            
            for (i = 0; i < rowLimit; i++)
            {
                var row = sheet.CreateRow(i);
                var cell1 = row.CreateCell(0);
                var cell2 = row.CreateCell(1);
                var cell3 = row.CreateCell(2);
                var cell4 = row.CreateCell(3);
                var cell5 = row.CreateCell(4);
                var cell6 = row.CreateCell(5);
                var cell7 = row.CreateCell(6);
                var cell8 = row.CreateCell(7);
                var cell9 = row.CreateCell(8);
                var cell10 = row.CreateCell(9);
                cell1.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell2.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell3.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell4.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell5.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell6.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell7.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell8.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell9.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
                cell10.SetCellValue("ZOMG PLEASE SURVIVE THIS STRESS TEST");
            }

            String file = @"C:\Users\Asus\OneDrive\Desktop\PruebasExcel\test.xls";
            using (FileStream fs = new FileStream(file, FileMode.Create))
            {
                workbook.Write(fs);
            }

            Console.WriteLine("Finalizado con libreria");
            System.Console.WriteLine("tiempo tomado libreria: " + (DateTime.Now - originalTime));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            #region prueba excel
            //writeToExcel();
            //Test();

            //var openFile = new OpenFileDialog()
            //{
            //    Filter = "Excel Files (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm",
            //    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            //    DefaultExt = "xlsx",
            //};
            //var res = openFile.ShowDialog();
            //if (res != true) return;

            //var fileName = openFile.FileName;
            //var woorkbook = ExcelControl.ReadWoorkBook(fileName);

            //var sheet = woorkbook.GetSheetAt(0);
            //var test = ExcelControl.GetValue(sheet, 0, 0);

            //MessageBox.Show(test.ToString());

            //ExcelControl.SetValue(sheet, 1, 1, 9999);

            //ExcelControl.SaveOrUptateWoorkBook(fileName, woorkbook);

            //DateTime originalTime = DateTime.Now;
            //System.Console.WriteLine("Busqueda iniciada: " + originalTime);
            //var searchRes = ExcelControl.SearchValue(sheet, "ESTA SI NO LA ENCUENTRAS");
            //System.Console.WriteLine("Tiempo de busqueda: " + (DateTime.Now - originalTime));

            //MessageBox.Show(searchRes.ToString());

            //var test = ExcelControl.GetValue(sheet, 1, 1);
            //var test2 = ExcelControl.GetValue(sheet, 1, 2);

            //MessageBox.Show(test.ToString());
            //MessageBox.Show(test2.ToString());
            #endregion

            TestWord();
        }

        public void TestWord()
        {
            XWPFDocument doc = new XWPFDocument();

            var datos = DatosPruebav3();
            var cols = GetNumeroColumnas(datos);
            var rows = datos.Count;


            XWPFTable table1 = doc.CreateTable(rows, cols);
            
            var tblLayout1 = table1.GetCTTbl().tblPr.AddNewTblLayout();


            tblLayout1.type = ST_TblLayoutType.autofit;

            table1.SetCellMargins(0, 0, 0, 0);

            

            foreach (var item in datos)
            {

            }



            using (FileStream fs = new FileStream(@"C:\Users\Asus\OneDrive\Desktop\PruebasExcel\complexTable.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }

        public int GetNumeroColumnas(List<string[]> DatosTabla)
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
