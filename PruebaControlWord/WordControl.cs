using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using static NPOI.XWPF.UserModel.XWPFTable;

namespace PruebaControlWord
{
    public class WordControl
    {
        /// <summary>
        /// Lee el archivo excel que se encuentra en la ruta dada como parametro
        /// </summary>
        /// <param name="path">Ruta del archivo</param>
        /// <returns>Documento de excel</returns>
        public static XWPFDocument ReadDocument(string path)
        {
            XWPFDocument book;
            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                book = new XWPFDocument(fs);

                fs.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); book = null; }

            return book;
        }

        /// <summary>
        /// Guarda los cambios realizados a un documento en una ruta dada
        /// </summary>
        /// <param name="path">Ruta de guardado</param>
        /// <param name="workbook">Documento de excel a guardar</param>
        /// <returns>Estado de guardado: true = guardado, false = error</returns>
        public static bool WriteDocument(string path, XWPFDocument workbook)
        {
            bool state;

            try
            {
                FileStream fs = new FileStream(path, FileMode.Create);
                workbook.Write(fs);
                state = true;

                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                state = false;
            }

            return state;
        }

        /// <summary>
        /// Crea un nuevo parrafo en el documento dado
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="text"></param>
        /// <param name="fontFamily"></param>
        /// <param name="fontSize"></param>
        /// <param name="alignment"></param>
        public static void CreateParagraph(XWPFDocument doc, string text, string fontFamily = "Arial", int fontSize = 12, ParagraphAlignment alignment = ParagraphAlignment.LEFT, CT_SectPr seccion = null)
        {
            if (seccion == null)
            {
                seccion = doc.Document.body.sectPr;
            }

            //doc.Document.body.sectPr = seccion;

            XWPFParagraph paragraph = doc.CreateParagraph();
            paragraph.SpacingAfter = 0;
            paragraph.SpacingBefore = 0;
            paragraph.Alignment = alignment;
            //paragraph.GetCTP().AddNewPPr().sectPr = seccion;

            XWPFRun run = paragraph.CreateRun();
            run.SetText(text);
            run.FontFamily = fontFamily;
            run.FontSize = fontSize;
        }

        /// <summary>
        /// Crea una tabla en el documento dado
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="DatosTabla"></param>
        /// <param name="Bordes"></param>
        /// <param name="Dobletitulo"></param>
        public static void CreateTable(XWPFDocument doc, List<string[]> DatosTabla, bool Bordes = false, bool Dobletitulo = false, CT_SectPr seccion = null)
        {
            if (seccion == null) seccion = doc.Document.body.sectPr;

            var pagewidth = seccion.pgSz.w;
            var pagemargw = seccion.pgMar.left + seccion.pgMar.right;
            var effectivewidth = pagewidth - pagemargw;
            

            var cols = GetNumeroColumnas(DatosTabla);
            var table = doc.CreateTable(DatosTabla.Count, cols);
            
            var tablepr = table.GetCTTbl().AddNewTblPr();
            tablepr.tblW = new CT_TblWidth();
            tablepr.tblW.type = ST_TblWidth.dxa;
            tablepr.tblW.w = effectivewidth.ToString();
            
            table.SetCellMargins(0, 0, 0, 0);

            for (int rowN = 0; rowN < DatosTabla.Count; rowN++)
            { 
                var rowElements = DatosTabla[rowN];
                var rowTable = table.GetRow(rowN);


                if (rowElements.Length < cols && rowElements.Length == 1)
                {
                    rowTable.MergeCells(0, cols - 1);

                    var titleCell = rowTable.GetCell(0);
                    titleCell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    titleCell.SetParagraph(SetCellText(table, rowElements[0], true, 12, "Arial", ParagraphAlignment.CENTER, Bordes ? "#FFFFFF" : "#000000"));
                    SetCellBorder(titleCell, Bordes);
                    
                    if (Bordes) titleCell.SetColor("#FF0000");
                }
                else
                {
                    for (int colN = 0; colN < rowElements.Length; colN++)
                    {
                        XWPFTableCell cellEdit = rowTable.GetCell(colN);

                        CT_TcPr m_Pr = cellEdit.GetCTTc().AddNewTcPr();

                        var esdobletitulo = Dobletitulo && rowN == 1;
                        var last = rowElements[colN] != "" ? rowElements[colN].Last() : 'a';
                        var text = last.Equals('¬') ? rowElements[colN].Remove(rowElements[colN].Length - 1, 1) : rowElements[colN];
                        var size = esdobletitulo ? 12 : 11;
                        var align = last.Equals('¬') ? ParagraphAlignment.LEFT : ParagraphAlignment.CENTER;
                        var color = esdobletitulo ? "#FFFFFF" : "#000000";

                        cellEdit.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        
                        cellEdit.SetParagraph(SetCellText(table, text, esdobletitulo, size, "Arial", align, color));
                        SetCellBorder(cellEdit, Bordes);
                        if (esdobletitulo) cellEdit.SetColor("#FF0000");
                    }
                }
            }
        }

        public static XWPFParagraph SetCellText(XWPFTable table, string text, bool isBold = false, int fontSize = 11, string fontFamily = "Arial", ParagraphAlignment alignment = ParagraphAlignment.CENTER, string color = "#000000")
        {
            CT_P para = new CT_P();
            XWPFParagraph pCell = new XWPFParagraph(para, table.Body);
            
            pCell.Alignment = alignment;
            pCell.VerticalAlignment = NPOI.XWPF.UserModel.TextAlignment.CENTER;
            pCell.SpacingBefore = 0;
            pCell.SpacingAfter = 0;

            XWPFRun r1c1 = pCell.CreateRun();
            r1c1.SetText(text);
            r1c1.FontSize = fontSize;
            r1c1.FontFamily = fontFamily;
            r1c1.IsBold = isBold;
            r1c1.SetColor(color);

            return pCell;
        }

        


        #region other
        public static int GetNumeroColumnas(List<string[]> DatosTabla)
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

        public static void SetCellBorder(XWPFTableCell cellEdit, bool Bordes)
        {
            if (Bordes)
            {
                cellEdit.SetBorderBottom(XWPFBorderType.SINGLE, 1, 0, "#000000");
                cellEdit.SetBorderTop(XWPFBorderType.SINGLE, 1, 0, "#000000");
                cellEdit.SetBorderLeft(XWPFBorderType.SINGLE, 1, 0, "#000000");
                cellEdit.SetBorderRight(XWPFBorderType.SINGLE, 1, 0, "#000000");
            }
            else
            {
                cellEdit.SetBorderBottom(XWPFBorderType.NONE, 0, 0, "#000000");
                cellEdit.SetBorderTop(XWPFBorderType.NONE, 0, 0, "#000000");
                cellEdit.SetBorderLeft(XWPFBorderType.NONE, 0, 0, "#000000");
                cellEdit.SetBorderRight(XWPFBorderType.NONE, 0, 0, "#000000");
            }
        }
        #endregion
    }
}
