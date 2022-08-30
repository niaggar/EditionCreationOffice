using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace PruebaControlOpenXML
{
    public class ImageControl
    {
        public static void CopyImage(string fileName, int length = 1)
        {
            string sourceFile = fileName;
            string destinationRoute = @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\images\";

            for (int i = 0; i < length; i++)
            {
                try
                {
                    File.Copy(sourceFile, destinationRoute + $"img-{i}.jpeg", true);
                }
                catch (IOException iox)
                {
                    Console.WriteLine(iox.Message);
                }
            }
        }
        
        public static string ImageToBase64(BitmapImage imageSource)
        {
            Stream stream = imageSource.StreamSource;
            Byte[] buffer = null;
            
            if (stream != null && stream.Length > 0)
            {
                using (BinaryReader br = new BinaryReader(stream))
                {
                    buffer = br.ReadBytes((Int32)stream.Length);
                }
            }

            var res = Convert.ToBase64String(buffer);
            System.Console.WriteLine(res);

            return res;
        }

        public static BitmapImage Base64StringToBitmap(string base64String)
        {
            byte[] imgBytes = Convert.FromBase64String(base64String);

            BitmapImage bitmapImage = new BitmapImage();
            MemoryStream ms = new MemoryStream(imgBytes);
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = ms;
            bitmapImage.EndInit();

            return bitmapImage;
        }


        public static void WriteTextToFile(string[] lines)
        {
            string destinationRoute = @"C:\Users\Asus\OneDrive\Desktop\PruebasOffice\images\img.txt";
            using (StreamWriter file = new StreamWriter(destinationRoute))
            {
                foreach (string line in lines)
                {
                    file.WriteLine(line);
                }
            }           
        }
    }
}
