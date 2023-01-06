using DIS.Reportes.Automatizados.DocTemplates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApps
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Template1.Create();

            // console hello
            Console.WriteLine("Hola");
            Console.ReadLine();
        }
    }
}
