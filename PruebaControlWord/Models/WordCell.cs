using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PruebaControlWord.Models
{
    public class WordCell
    {
        public int Width { get; set; }
        public string Content { get; set; }
        public int FontSize { get; set; }
        public string FontFamily { get; set; }
        public bool FontBold { get; set; }
        public ParagraphAlignment H_Alignment { get; set; }
        public TextAlignment V_Alignment { get; set; }
        private int mergeColumnNumber;
        public int MergeColumnNumber
        {
            get
            {
                if (mergeColumnNumber <= 0)
                    return 1;
                else
                    return mergeColumnNumber;
            }
            set
            {
                mergeColumnNumber = value;
            }
        }

        

        public WordCell()
        {
            FontSize = 10;
            FontFamily = "Arial";
            FontBold = false;
            H_Alignment = ParagraphAlignment.LEFT;
            V_Alignment = TextAlignment.CENTER;
            Content = string.Empty;
        }
    }
}
