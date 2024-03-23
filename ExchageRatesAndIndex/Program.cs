using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExchageRatesAndIndex
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            
            clsMain obj = new clsMain();
            //System.Windows.Forms.Application.Run();
            Environment.Exit(0);
        }
    }
}
