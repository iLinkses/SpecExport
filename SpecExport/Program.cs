using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using SpecExport.Classes;
using ConsoleTables;

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
                WConsoleTable(
                    new string[] { "Project", "Author", "Vesrion" },
                    new List<object[]> {
                        new object[] {
                            System.Reflection.Assembly.GetExecutingAssembly().GetName().Name,
                            "iLinks",
                            System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
                        } }
                    );

                log = LogManager.GetCurrentClassLogger();
                log.Trace($"---!!!---Start NLog---!!!---\n" +
                    $"\tMachineName:{Environment.MachineName}\n" +
                    $"\tUserName:{Environment.UserName}\n" +
                    $"\tOS:{Environment.OSVersion.VersionString}");

                WConsoleTable(
                    new string[] { "MachineName", "UserName", "OSVersion" },
                    new List<object[]>
                    {
                        new object[] {
                            Environment.MachineName,
                            Environment.UserName,
                            Environment.OSVersion.VersionString
                        } }
                    );

                //Kompas kompas = new Kompas();
                //kompas.ExportSpec();
                GetTestData();

                if (Properties.Settings.Default.SendMail)
                {
                    Classes.SMTP smtp = new Classes.SMTP();
                    smtp.SendMail();
                }

                log.Trace("---!!!---End NLog---!!!---");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка работы с логом!\n{ex.Message}");
                Console.ReadLine();
            }
        }
        public static void WConsoleTable(string[] Columns, List<object[]> Rows)
        {
            var ctbl = new ConsoleTable();
            ctbl.AddColumn(Columns);
            foreach (var row in Rows)
            {
                ctbl.AddRow(row);
            }
            ctbl.Options.EnableCount = false;
            ctbl.Write();
        }
        private static void GetTestData()
        {
            Spec spec = new Spec();
            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Болт 1",
                Name = "Болт 1",
                Quantity = 1,
                Section = "Лист1"
            });
            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Болт 2",
                Name = "Болт 2",
                Quantity = 2,
                Section = "Лист1"
            });
            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Шайба 1",
                Name = "Шайба 1",
                Quantity = 2,
                Section = "Лист1"
            });
            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 1",
                Name = "Гайка 1",
                Quantity = 1,
                Section = "Лист1"
            });

            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 1",
                Name = "Гайка 1",
                Quantity = 1,
                Section = "Лист2"
            });
            spec.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 2",
                Name = "Гайка 2",
                Quantity = 2,
                Section = "Лист2"
            });
            List<Spec> Specs = new List<Spec>();
            Specs.Add(new Spec
            {
                FileName = "1_",
                Positions = spec.Positions
            });
            ExcelExport excelExport = new ExcelExport(Specs);
            excelExport.SaveExcel();
        }
    }
}
