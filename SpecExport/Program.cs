using System;
using System.Collections.Generic;
using NLog;
using SpecExport.Classes;
using ConsoleTables;
using SpecExport.Properties;

namespace SpecExport
{
    class Program
    {
        ///<seealso cref="https://nlog-project.org/config/?tab=layout-renderers"/>
        public static Logger log { get; set; }
        /// <summary>
        /// Признак отображения таблиц в консоли
        /// </summary>
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

                Kompas kompas = new Kompas();
                kompas.ExportSpec();
                //GetTestData();

                if (Settings.Default.SendMail)
                {
                    SMTP smtp = new SMTP();
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
            if (Settings.Default.ShowConsoleTable)
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
            //---
            Spec spec1 = new Spec();
            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Болт 1",
                Name = "Болт 1",
                Quantity = 1,
                Section = "Лист1"
            });
            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Болт 3",
                Name = "Болт 3",
                Quantity = 2,
                Section = "Лист1"
            });
            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Шайба 2",
                Name = "Шайба 1",
                Quantity = 2,
                Section = "Лист1"
            });
            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 2",
                Name = "Гайка 1",
                Quantity = 1,
                Section = "Лист1"
            });

            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 1",
                Name = "Гайка 1",
                Quantity = 1,
                Section = "Лист2"
            });
            spec1.Positions.Add(new Spec.Detail
            {
                Designation = "Гайка 3",
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
            Specs.Add(new Spec
            {
                FileName = "2_",
                Positions = spec1.Positions
            });
            ExcelExport excelExport = new ExcelExport(Specs);
            excelExport.NewSaveExcel();
        }
    }
}
