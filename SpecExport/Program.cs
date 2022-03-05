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
        ///<seealso cref="https://nlog-project.org/config/?tab=layout-renderers"/>
        public static Logger log;
        static void Main(string[] args)
        {
            try
            {
                log = LogManager.GetCurrentClassLogger();
                log.Trace($"---!!!---Start NLog---!!!---\n" +
                    $"\tMachineName:{Environment.MachineName}\n" +
                    $"\tUserName:{Environment.UserName}\n" +
                    $"\tOS:{Environment.OSVersion.VersionString}");

                GetFileInCatalog();

                if (Properties.Settings.Default.SendMail)
                {
                    Classes.SMTP smtp = new Classes.SMTP();
                    smtp.SendMail();
                }

                log.Trace("---!!!---End NLog---!!!---");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка работы с логом!\n{ex.Message}");
                Console.ReadLine();
            }
        }

        static List<string> FileNames { get; set; } = new List<string>();
        /// <summary>
        /// Получает названия файлов с чертежами в каталоге
        /// </summary>
        private static void GetFileInCatalog()
        {
            string DrawingsDirectory = Properties.Settings.Default.DrawingsDirectory;
            foreach (var f in Directory.GetFiles(DrawingsDirectory, "*.dwg"))
            {
                FileNames.Add(Path.GetFileName(f));
                Console.WriteLine(Path.GetFileName(f));
            }
            if (FileNames.Count > 0) log.Trace($"Список чертежей получен: {FileNames}");
            else log.Error($"Пустой каталог {DrawingsDirectory}");
        }
    }
}
