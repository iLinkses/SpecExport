using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using SpecExport.Classes;

namespace SpecExport
{
    class Program
    {
        ///<seealso cref="https://nlog-project.org/config/?tab=layout-renderers"/>
        public static Logger log { get; set; }
        static void Main(string[] args)
        {
            try
            {
                log = LogManager.GetCurrentClassLogger();
                log.Trace($"---!!!---Start NLog---!!!---\n" +
                    $"\tMachineName:{Environment.MachineName}\n" +
                    $"\tUserName:{Environment.UserName}\n" +
                    $"\tOS:{Environment.OSVersion.VersionString}");
                Kompas kompas = new Kompas();
                kompas.ExportSpec();

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
    }
}
