using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace SpecExport
{
    class Program
    {
        public static Logger log;
        static void Main(string[] args)
        {
            StartLogging();
        }

        static void StartLogging()
        {
            try
            {
                log = LogManager.GetCurrentClassLogger();

                log.Trace("MachineName: {0}", Environment.MachineName);
                log.Trace("UserName: {0}", Environment.UserName);
                log.Trace("OS: {0}", Environment.OSVersion.ToString());
                log.Trace("Command: {0}", Environment.CommandLine.ToString());

                //NLog.Targets.FileTarget tar = (NLog.Targets.FileTarget)LogManager.Configuration.FindTargetByName("filedata");
                //tar.DeleteOldFileOnStartup = false;
                GetFileInCatalog();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Ошибка работы с логом!\n{e.Message}");
            }
            Console.ReadLine();
        }

        static List<string> FileNames { get; set; } = new List<string>();
        /// <summary>
        /// Получает названия файлов с чертежами в каталоге
        /// </summary>
        private static void GetFileInCatalog()
        {
            string DrawingsDirectory = "Drawings";
            foreach (var f in Directory.GetFiles(DrawingsDirectory, "*.dwg"))
            {
                FileNames.Add(Path.GetFileName(f));
                Console.WriteLine(Path.GetFileName(f));
            }
        }
    }
}
