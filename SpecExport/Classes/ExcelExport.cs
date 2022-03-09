using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;

namespace SpecExport.Classes
{
    class ExcelExport
    {
        private readonly NLog.Logger log = Program.log;
        private List<Spec> Specs { get; set; } = new List<Spec>();
        private string DrawingsDirectory { get { return Properties.Settings.Default.DrawingsDirectory; } }
        private readonly string NameDoc = $"Отчет за {DateTime.Now.ToShortDateString().Replace(".", "_")}.xlsx";
        //private Spec spec;

        public ExcelExport(List<Spec> spec_)
        {
            this.Specs = spec_;
            log = LogManager.GetCurrentClassLogger();
        }
        public void SaveExcel()
        {
            try
            {
                if (!ExistsDrawingsDirectory())
                {
                    Directory.CreateDirectory(DrawingsDirectory);
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                if (!ExistsFile())
                {
                    //создаем новый файл
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        foreach (var spec in Specs)
                        {
                            foreach (var sec in spec.Positions.Select(p => p.Section).Distinct())
                            {
                                if (ep.Workbook.Worksheets.Select(ws_ => ws_.Name).ToList().Find(n => n == sec) == null)
                                {
                                    ExcelWorksheet ws = ep.Workbook.Worksheets.Add(sec);
                                    ///Названия товаров
                                    foreach (var detail in spec.Positions.Where(p => p.Section.Equals(sec)))
                                    {
                                        ws.Cells[1, 1].Value = "№";
                                        ws.Cells[1, 2].Value = "Название";
                                        ws.Cells[1, 3].Value = "Общее кол-во";
                                    }
                                }
                            }
                        }
                        ep.SaveAs($@"{DrawingsDirectory}\{NameDoc}");
                        Console.WriteLine($"Создан новый пустой файл {NameDoc}");
                        log.Trace($"Создан новый пустой файл {NameDoc}");
                    }
                }
                using (ExcelPackage ep = new ExcelPackage(new FileInfo($@"{DrawingsDirectory}\{NameDoc}")))
                {
                    foreach (var ws in ep.Workbook.Worksheets)
                    {
                        //Console.WriteLine(ws.Name);
                        foreach (var spec in Specs)
                        {
                            AddSubsection(ws, spec);
                            //AddCells(ws, spec);

                            //foreach (var test in ws.Cells[ws.Dimension.FullAddress])
                            //{
                            //    Console.WriteLine(test.Value);
                            //}
                        }
                        //var i = $"C2:C{ws.Dimension.End.Row}";
                        for (int row = 2; row <= ws.Dimension.End.Row; row++)
                        {
                            //Добавляем формулу для расчета общей суммы
                            ws.Cells[row, 3].Formula = $"=SUM(D{row}:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}{row})";
                        }

                        string[] columns = new string[ws.Dimension.End.Column];
                        for (int col = 1; col <= ws.Dimension.End.Column; col++)
                        {
                            columns[col - 1] = ws.Cells[1, col].Value.ToString();
                        }
                        List<object[]> crs = new List<object[]>();
                        for (int row = 2; row <= ws.Dimension.End.Row; row++)
                        {
                            object[] cr = new object[ws.Dimension.End.Column];
                            for (int col = 1; col <= ws.Dimension.End.Column; col++)
                            {
                                cr[col - 1] = ws.Cells[row, col].Value;
                            }
                            crs.Add(cr);
                        }
                        Program.WConsoleTable(columns, crs);

                        ws.Cells.AutoFitColumns();
                    }
                    ep.Save();
                    Console.WriteLine($"{NameDoc} сгенерирован!");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, ex.Message);
                Console.WriteLine("Возникла ошибка. Подробности см. в логе");
            }
        }

        private void AddSubsection(ExcelWorksheet ws, Spec spec)
        {
            int row = 2, col = 1;
            foreach(var ssn in spec.Positions.Where(p => p.Section == ws.Name).Select(n => Regex.Match(n.Name, @"[а-яА-Я]+").Value).Distinct())
            {
                ws.Cells[row, col].Value = ssn;
                row += 2;
            }
            //TODO Создать метод возвращающий первую и последнюю строку в подразделе
            //TODO Объединить 4 колонки в одну для названия подраздела
            //TODO както распределять позиции по подразделам (тоже метод написать)
        }

        private void AddCells(ExcelWorksheet ws, Spec spec)
        {
            int row = 2;
            int col = 4;
            try
            {
                //Проверка на наличие добавленных чертежей
                if (ws.Cells["D1"].Value == null)
                {
                    ws.SetValue("D1", spec.DrawingNumber);
                }
                else
                {
                    var test = GetLastNotNullCol(ws) + 1;
                    var test1 = ws.Cells[1, GetLastNotNullCol(ws) + 1].Value;
                    ws.Cells[1, GetLastNotNullCol(ws) + 1].Value = spec.DrawingNumber;//Добавляем в столбец номер очередного чертежа
                }

                foreach (var p in spec.Positions.Where(p => p.Section == ws.Name))
                {
                    //Проверка на наличие записей впринципе
                    if (ws.Cells["B2"].Value == null)
                    {
                        //Не нашли записей, добавляем все подряд
                        ws.Cells[row, 2].Value = p.Name;//Добавляем название
                        ws.Cells[row, col].Value = p.Quantity;//Добавляем количество
                        row++;
                    }
                    else
                    {
                        string existRow = ExistName(ws.Cells[$"B2:B{ws.Dimension.End.Row}"], p.Name);
                        //Делаем поиск на существующее название
                        if (string.IsNullOrEmpty(existRow))
                        {
                            //Не нашли одинакового, добавляем в конец
                            row = GetLastNotNullRow(ws) + 1;
                            ws.SetValue(row, 2, p.Name);
                            string existCol = ExistName(ws.Cells[$"D1:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}1"], spec.DrawingNumber.ToString());
                            if (string.IsNullOrEmpty(existCol))
                            {
                                //Не нашли такой номер добавляем в конец
                                ws.SetValue(row, GetLastNotNullCol(ws) + 1, p.Quantity);
                            }
                            else
                            {
                                //Нашли такой же номер
                                ws.SetValue(new string(existCol.ToCharArray().Where(n => !char.IsDigit(n)).ToArray()) + row, p.Quantity);
                            }
                        }
                        else
                        {
                            //Нашли такое же
                            string existCol = ExistName(ws.Cells[$"D1:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}1"], spec.DrawingNumber.ToString());
                            if (string.IsNullOrEmpty(existCol))
                            {
                                //Не нашли такой номер добавляем в конец
                                ws.SetValue(Convert.ToInt32(new string(existRow.ToCharArray().Where(n => char.IsDigit(n)).ToArray())), GetLastNotNullCol(ws) + 1, p.Quantity);
                            }
                            else
                            {
                                //Нашли такой же номер
                                ws.SetValue(new string(existCol.ToCharArray().Where(n => !char.IsDigit(n)).ToArray()) + new string(existRow.ToCharArray().Where(n => char.IsDigit(n)).ToArray()), p.Quantity);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, ex.Message);
                Console.WriteLine("Возникла ошибка. Подробности см. в логе");
            }
        }
        /// <summary>
        /// Поиск последней not null строки
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        private int GetLastNotNullRow(ExcelWorksheet ws)
        {
            int row_ = 2;
            for (int row = 2; row <= ws.Cells.End.Row; row++)
            {
                if (ws.GetValue(row, 2) == null && ws.GetValue(row - 1, 2) != null)
                {
                    row_ = row - 1;
                    break;
                }
            }
            return row_;
        }
        /// <summary>
        /// Поиск последнего not null столбца
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        private int GetLastNotNullCol(ExcelWorksheet ws)
        {
            int col_ = 4;
            for (int col = 4; col <= ws.Cells.End.Column; col++)
            {
                if (ws.GetValue(1, col) == null && ws.GetValue(1, col - 1) != null)
                {
                    col_ = col - 1;
                    break;
                }
            }
            return col_;
        }
        /// <summary>
        /// Проверяет существует ли дефолтная директория
        /// </summary>
        /// <returns></returns>
        private bool ExistsDrawingsDirectory()
        {
            if (Directory.Exists(DrawingsDirectory)) return true;
            else return false;
        }
        /// <summary>
        /// Проверяет существует ли файл
        /// </summary>
        /// <returns></returns>
        private bool ExistsFile()
        {
            if (File.Exists($@"{DrawingsDirectory}\{NameDoc}")) return true;
            else return false;
        }
        /// <summary>
        /// Проверяет была ли уже добавлена деталь с таким именем
        /// </summary>
        /// <param name="er"></param>
        /// <param name="name"></param>
        /// <returns>Адрес ячейки куда добавлено название (мы потом из нее выкусывает номер строки)</returns>
        private string ExistName(ExcelRange er, string name)
        {
            string existNameAdress = string.Empty;
            foreach (var t in er)
            {
                if (t.Value.ToString() == name)
                {
                    existNameAdress = t.Address;
                    break;
                }
            }
            return existNameAdress;
        }
        private void AddThinBorder(ExcelWorksheet ws, int RowStart, int ColumnStart, int RowEnd, int ColumnEnd)
        {
            ws.Cells[RowStart, ColumnStart, RowEnd, ColumnEnd].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[RowStart, ColumnStart, RowEnd, ColumnEnd].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[RowStart, ColumnStart, RowEnd, ColumnEnd].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells[RowStart, ColumnStart, RowEnd, ColumnEnd].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        private void AddThickBorder(ExcelWorksheet ws, int RowStart, int ColumnStart, int RowEnd, int ColumnEnd)
        {
            ws.Cells[RowStart, ColumnStart, RowStart, ColumnEnd].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells[RowStart, ColumnEnd, RowEnd, ColumnEnd].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells[RowEnd, ColumnStart, RowEnd, ColumnEnd].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells[RowStart, ColumnStart, RowEnd, ColumnStart].Style.Border.Left.Style = ExcelBorderStyle.Thick;
        }
        private void SetAutoFitColumns(ExcelWorksheet ws, int RowStart, int ColumnStart, int RowEnd, int ColumnEnd)
        {
            ws.Cells[RowStart, ColumnStart, RowStart, ColumnEnd].AutoFitColumns();
        }
    }
}
