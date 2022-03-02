using System;
using System.Collections.Generic;
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
            }
            catch (Exception e)
            {
                Console.WriteLine($"Ошибка работы с логом!\n{e.Message}");
            }
            Console.ReadLine();
        }
    }
}
