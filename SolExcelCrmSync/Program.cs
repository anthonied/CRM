using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SolExcelCrmSync.Classes;

namespace SolExcelCrmSync
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);                    
            
            if(args.Length > 0 )
            {
                if (args[0] == "-Talisman")
                {
                    Application.Run(new TalismanImport());
                }
            }
            else
            {
            Application.Run(new Automation(args));
            }
        }
    }
}
