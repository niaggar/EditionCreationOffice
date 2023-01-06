using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIS.Reportes.Automatizados.Utils
{
    public class UnitsConverter
    {
        /// <summary>
        /// Convertir pulgada a twip
        /// </summary>
        /// <param name="valcm"></param>
        /// <returns></returns>
        public static double ConvertInchToTwip(double valp)
        {
            return valp * 1440;
        }

        /// <summary>
        /// Convertir cm a twip
        /// </summary>
        /// <param name="valcm"></param>
        /// <returns></returns>
        public static double ConvertCmToTwip(double valcm)
        {
            return valcm * 567;
        }

        /// <summary>
        /// Convertir twip a cm
        /// </summary>
        /// <param name="valtwip"></param>
        /// <returns></returns>
        public static double ConvertTwipToCm(double valtwip)
        {
            return valtwip * (1 / 567);
        }

        /// <summary>
        /// Convertir twip a cm
        /// </summary>
        /// <param name="valtwip"></param>
        /// <returns></returns>
        public static double ConvertTwipToInch(double valtwip)
        {
            return valtwip * (1 / 1440);
        }
    }
}
