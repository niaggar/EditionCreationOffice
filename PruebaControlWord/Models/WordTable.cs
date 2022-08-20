using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PruebaControlWord.Models
{
    public class WordTable
    {
        public int ColumnCount { get; set; }
        public int Width { get; set; }
        public List<int> ColumnWidth { get; set; }
        public List<List<WordCell>> Rows { get; set; }
        public List<WordCell> CurrentRow
        {
            get
            {
                if (Rows == null)
                    throw new Exception("row index out of range");

                var currentRow = Rows.LastOrDefault();

                return currentRow;
            }
        }

        
        public WordTable(int columnCount)
        {
            this.ColumnCount = columnCount;
        }

        public WordTable(int columnCount, List<List<WordCell>> rows)
        {
            this.ColumnCount = columnCount;
            this.Rows = rows;
        }
        
       

        public List<WordCell> AddRow()
        {
            if (this.Rows == null)
            {
                Rows = new List<List<WordCell>>();
            }

            var newRow = new List<WordCell>();
            for (var i = 0; i < ColumnCount; i++)
            {
                var newCell = new WordCell();
                newRow.Add(newCell);
            }

            Rows.Add(newRow);

            return newRow;
        }

        public WordCell GetCell(int row, int col)
        {
            if (Rows == null || Rows.Count < row + 1)
                return null;

            var currentRow = Rows[row];

            if (currentRow == null || currentRow.Count < col + 1)
                return null;

            return currentRow[col];
        }

        public List<WordCell> GetRow(int row)
        {
            if (Rows == null || Rows.Count < row + 1)
                return null;

            var currentRow = Rows[row];

            return currentRow;
        }

        public void Merge(int row, int col, int mergeNumber)
        {
            var currentRow = GetRow(row);

            if (currentRow == null)
                throw new Exception("row index out of range");

            if (col + 1 + mergeNumber > ColumnCount)
                throw new Exception("column index out of range");

            for (var i = col + 1; i < col + 1 + mergeNumber; i++)
            {
                var needRemoveCell = GetCell(row, i);

                if (needRemoveCell != null)
                    currentRow.Remove(needRemoveCell);
            }
        }

        public void MergeCurrentRow(int col, int mergeNumber)
        {
            if (CurrentRow == null)
                throw new Exception("row index out of range");

            if (col + 1 + mergeNumber > ColumnCount)
                throw new Exception("column index out of range");

            CurrentRow[col].MergeColumnNumber = mergeNumber + 1;

            for (var i = col + 1; i < col + 1 + mergeNumber; i++)
            {
                var needRemoveCell = CurrentRow[i];

                if (needRemoveCell != null)
                    CurrentRow.Remove(needRemoveCell);
            }
        }

        public void AdjustColumnWidth()
        {
            if (ColumnWidth == null || ColumnWidth.Count != ColumnCount || Rows == null || Rows.Count == 0)
                throw new Exception("column index out of range");

            for (var i = 0; i < ColumnWidth.Count; i++)
            {
                var w = ColumnWidth[i];

                foreach (var row in Rows)
                {
                    if (row == null || row.Count == 0)
                        continue;

                    for (var c = 0; c < row.Count; c++)
                    {
                        if (c == i)
                        {
                            row[c].Width = w;
                        }
                    }
                }
            }
        }
    }
}
