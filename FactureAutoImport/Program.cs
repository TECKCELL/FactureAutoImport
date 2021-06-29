using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace FactureAutoImport
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        static void Main(string[] args)
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new AutoImport()
            };
            ServiceBase.Run(ServicesToRun);
           /* if (Environment.UserInteractive)
            {
                AutoImport service1 = new AutoImport();
                service1.TestStartupAndStop(args);
            }*/
        }
    }
}
