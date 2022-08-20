using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.XSSF.UserModel; // .xls
using NPOI.HSSF.UserModel; // .xlsx
using NPOI.SS.UserModel;   // .xls y .xlsx
using System.IO;
using System.Windows;
using NPOI.SS.Util;

namespace PruebaControlExcel
{
    public class ExcelControl
    {
        /// <summary>
        /// Obtener el nombre de las hojas presentes en el excel
        /// </summary>
        /// <param name="workbook">Documento de excel</param>
        /// <returns>Lista de nombres</returns>
        public static List<string> GetSheetsName(IWorkbook workbook)
        {
            List<string> sheets = new List<string>();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                sheets.Add(workbook.GetSheetName(i));
            }
            
            return sheets;
        }
        
        /// <summary>
        /// Lee el archivo excel que se encuentra en la ruta dada como parametro
        /// </summary>
        /// <param name="path">Ruta del archivo</param>
        /// <returns>Documento de excel</returns>
        public static IWorkbook ReadWoorkBook(string path)
        {
            IWorkbook book = null;

            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                // Try to read workbook as XLSX:
                try { book = new XSSFWorkbook(fs); }
                catch { book = null; }

                // If reading fails, try to read workbook as XLS:
                if (book == null) book = new HSSFWorkbook(fs);

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
        public static bool SaveOrUptateWoorkBook(string path, IWorkbook workbook)
        {
            bool state = false;

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
        /// Establece el valor para una celda en especifico
        /// </summary>
        /// <param name="sheet">Hoja de calculo que se va a editar</param>
        /// <param name="rowIndex">Indice de la fila</param>
        /// <param name="colIndex">Indice de la columna</param>
        /// <param name="value">Valor a establecer</param>
        public static void SetValue(ISheet sheet, int rowIndex, int colIndex, double value)
        {
            // Selecciona toda una fila
            IRow row = sheet.GetRow(rowIndex); 
            if (row == null) row = sheet.CreateRow(rowIndex);

            // Dentro de la fila selecciona una columna
            ICell cell = row.GetCell(colIndex);
            if (cell == null) cell = row.CreateCell(colIndex);

            // La celda comprendida entre la fila y columna dada se establece el valor
            cell.SetCellValue(value);
        }

        /// <summary>
        /// Busca la primera celda en contener el valor especificicado
        /// </summary>
        /// <param name="sheet">Hoja de calculo donde se busca el valor</param>
        /// <param name="valueSearch">Valor a buscar</param>
        /// <returns>Tupla = (fila de la celda, columna de la celda) - Error = (-1, -1)</returns>
        public static (int row, int col) SearchValue(ISheet sheet, object valueSearch)
        {
            int rowStart = sheet.FirstRowNum;
            int rowFinal = sheet.LastRowNum;
            if (rowStart == -1 || rowFinal == -1) return (-1, -1);

            for (int i = rowStart; i <= rowFinal; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                
                int colStart = row.FirstCellNum;
                int colFinal = row.LastCellNum;
                if (colStart == -1 || colFinal == -1) continue;

                for (int j = colStart; j <= colFinal; j++)
                {
                    var res = GetValue(sheet, i, j);

                    if (res.cellValue != null && res.cellValue.Equals(valueSearch))
                    {
                        return (i, j);
                    }
                }
            }

            return (-1, -1);
        }

        /// <summary>
        /// Obtener un valor de una celda, conociendo los inidices de su fila y columna
        /// </summary>
        /// <param name="sheet">Hoja de caalculo donde se encuentra el valor</param>
        /// <param name="rowIndex">Indice de la fila</param>
        /// <param name="colIndex">Indice de la columna</param>
        /// <returns>Tupla = (valor de la celda, tipo de la celda) - Error = (null, null)</returns>
        public static (object cellValue, Type cellType) GetValue(ISheet sheet, int rowIndex, int colIndex)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null) return (null, null);

            ICell cell = row.GetCell(colIndex);
            if (cell == null) return (null, null);

            object valToRes = GetCellValue(cell);

            return (valToRes, valToRes.GetType());
        }

        /// <summary>
        /// Obtiene el valor que contiene la celda dada como parametro
        /// </summary>
        /// <param name="cell">Celda a obtener valor</param>
        /// <returns>Valor de la celda, como objeto generico</returns>
        private static object GetCellValue(ICell cell)
        {
            object cValue;
            switch (cell.CellType)
            {
                case (CellType.Unknown | CellType.Formula | CellType.Blank):
                    cValue = cell.ToString();
                    break;
                case CellType.Numeric:
                    cValue = cell.NumericCellValue;
                    break;
                case CellType.String:
                    cValue = cell.StringCellValue;
                    break;
                case CellType.Boolean:
                    cValue = cell.BooleanCellValue;
                    break;
                case CellType.Error:
                    cValue = cell.ErrorCellValue;
                    break;
                default:
                    cValue = string.Empty;
                    break;
            }
            
            return cValue;
        }
    }
}
