using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace DIS.Reportes.Automatizados.Utils
{
    public class DocumentController
    {
        public static WordprocessingDocument Create(string route, bool autosave = true)
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

        public static WordprocessingDocument Open(string route)
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

        public static void SavePDF(string docFullPath, string pdfFullPath)
        {
            Application appWord = new Application();
            appWord.Visible = false;
            var wordDocument = appWord.Documents.Open(docFullPath);

            wordDocument.SaveAs2(pdfFullPath, WdSaveFormat.wdFormatPDF);
            wordDocument.Close(false);
            appWord.Quit();
        }
    }
}
