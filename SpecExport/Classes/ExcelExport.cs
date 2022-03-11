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
        private ExcelSpec ExcelSpec = new ExcelSpec();
        private string DrawingsDirectory { get { return Properties.Settings.Default.DrawingsDirectory; } }
        private readonly string NameDoc = $"Отчет за {DateTime.Now.ToShortDateString().Replace(".", "_")}.xlsx";
        //private Spec spec;
        private readonly List<string> Subsections = new List<string>() { "", "", "" };//TODO сюда записать стандартные подразделы
        private List<Tuple<string, int, int>> SubsectionsAddress { get; set; } = new List<Tuple<string, int, int>>();

        public ExcelExport(List<Spec> spec_)
        {
            this.Specs = spec_;
            log = LogManager.GetCurrentClassLogger();
            NewStruct();
        }
        private void NewStruct()
        {
            foreach (var spec in Specs)
            {
                foreach (var detail in spec.Positions)
                {
                    var ed = new ExcelSpec.ExcelDetail();
                    ed.FileName = spec.FileName;
                    ed.DrawingNumber = spec.DrawingNumber;
                    ed.Section = detail.Section;
                    ed.Subsection = detail.Subsection;
                    ed.Format = detail.Format;
                    ed.Zone = detail.Zone;
                    ed.Position = detail.Position;
                    ed.Designation = detail.Designation;
                    ed.Name = detail.Name;
                    ed.Quantity = detail.Quantity;
                    ed.Note = detail.Note;
                    ExcelSpec.Details.Add(ed);
                }
            }
            //var sections = ExcelSpec.Details.Select(d => d.Section).Distinct().ToList();
            //var subsections = ExcelSpec.Details.Select(d => d.Subsection).Distinct().ToList();
        }

        public void NewSaveExcel()
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
                        foreach (var sec in ExcelSpec.Details.Select(d => d.Section).Distinct())
                        {
                            ExcelWorksheet ws = ep.Workbook.Worksheets.Add(sec);
                            ws.Cells[1, 1].Value = "№";
                            ws.Cells[1, 2].Value = "Название";
                            ws.Cells[1, 3].Value = "Общее кол-во";
                            ws.View.FreezePanes(2, 1);
                            int row = 2, col = 4;
                            var tsubsec = ExcelSpec.Details.Where(d => d.Section == sec).Select(d => d.Subsection).Distinct();
                            foreach (var subsec in ExcelSpec.Details.Where(d => d.Section == sec).Select(d => d.Subsection).Distinct())
                            {
                                ws.Cells[row, 1].Value = subsec;
                                row++;
                                var tdb = ExcelSpec.Details.Where(d => d.Section == sec && d.Subsection == subsec).Select(d => d.DrawingNumber).Distinct().OrderBy(d => d);
                                foreach (var dn in ExcelSpec.Details.Where(d => d.Section == sec && d.Subsection == subsec).Select(d => d.DrawingNumber).Distinct().OrderBy(d => d))
                                {
                                    if (ws.Cells["D1"].Value == null)
                                    {
                                        ws.SetValue("D1", dn);
                                    }
                                    else
                                    {
                                        string existCol = ExistName(ws.Cells[$"D1:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}1"], dn.ToString());
                                        if (string.IsNullOrEmpty(existCol))
                                        {
                                            ws.SetValue(1, col, dn);
                                        }
                                        else
                                        {
                                            col = ws.Cells[existCol].Start.Column;
                                        }
                                    }
                                    
                                    foreach (var detail in ExcelSpec.Details.Where(d => d.Section == sec && d.Subsection == subsec && d.DrawingNumber == dn).OrderBy(d => d.Name))
                                    {
                                        int row_ = row;
                                        string existRow = ExistName(ws.Cells[$"B2:B{ws.Dimension.End.Row}"],detail.Name);
                                        if (string.IsNullOrEmpty(existRow))
                                        {
                                            ws.SetValue(row, 2, detail.Name);
                                            row++;
                                        }
                                        else
                                        {
                                            row_ = ws.Cells[existRow].Start.Row;
                                        }
                                        ws.SetValue(row_, col, detail.Quantity);
                                    }
                                    col++;
                                }
                                
                            }
                            for (int r = 2; r <= ws.Dimension.End.Row; r++)
                            {
                                if (ws.Cells[r, 2].Value != null)
                                {
                                    ws.Cells[r, 3].Formula = $"=SUM(D{r}:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}{r})";
                                }
                                else
                                {
                                    ws.Cells[r, 1, r, 3].Merge = true;
                                    ws.Cells[r, 1, r, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }
                            }
                            ws.Cells.AutoFitColumns();
                            AddThinBorder(ws,ws.Dimension.Start.Row, ws.Dimension.Start.Column, ws.Dimension.End.Row, ws.Dimension.End.Column);

                            string[] columns = new string[ws.Dimension.End.Column];
                            for (int с = 1; с <= ws.Dimension.End.Column; с++)
                            {
                                columns[с - 1] = ws.Cells[1, с].Value.ToString();
                            }
                            List<object[]> crs = new List<object[]>();
                            for (int r = 2; r <= ws.Dimension.End.Row; r++)
                            {
                                object[] cr = new object[ws.Dimension.End.Column];
                                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                                {
                                    cr[c - 1] = ws.Cells[r, c].Value;
                                }
                                crs.Add(cr);
                            }
                            Program.WConsoleTable(columns, crs);
                        }
                        ep.SaveAs($@"{DrawingsDirectory}\{NameDoc}");
                        Console.WriteLine($"Отчет {NameDoc} сгенерирован!");
                        log.Trace($"Отчет {NameDoc} сгенерирован!");
                    }
                }
                else
                {
                    log.Error("Отчет уже был сгенерирован!");
                    Console.WriteLine("Отчет уже был сгенерирован!");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, ex.Message);
                Console.WriteLine("Возникла ошибка. Подробности см. в логе");
            }
        }
        [Obsolete("Первоначальная не оч реализация")]
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
                                    ws.View.FreezePanes(2, 1);///Закрепляет 1ю строку

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
                        Program.WConsoleTable(new string[] { ws.Name }, new List<object[]> { });
                        foreach (var spec in Specs)
                        {
                            //AddSubsection(ws, spec);
                            AddCells(ws, spec);

                            //foreach (var test in ws.Cells[ws.Dimension.FullAddress])
                            //{
                            //    Console.WriteLine(test.Value);
                            //}
                        }
                        //var i = $"C2:C{ws.Dimension.End.Row}";
                        for (int row = 2; row <= ws.Dimension.End.Row; row++)
                        {
                            var test = SubsectionsAddress.Exists(ssa => ssa.Item2 <= row || row <= ssa.Item3);
                            if (!SubsectionsAddress.Exists(ssa => ssa.Item2 == row))//SubsectionAddress.ContainsValue(row))
                            {
                                //Добавляем формулу для расчета общей суммы
                                ws.Cells[row, 3].Formula = $"=SUM(D{row}:{ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column)}{row})";
                            }
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
                        SubsectionsAddress.Clear();
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
            //TODO сделать список подразделов которые бывают стабильно, и сравнивать с подразделом из Detail.Subsection; если нашлось, добавлять в соответствующий раздел, если нет, то в "остальное/прочее"
            int row = 2, col = 1;
            foreach (var ssn in spec.Positions.Where(p => p.Section == ws.Name).Select(n => n.Subsection).Distinct())
            {
                SubsectionsAddress.Add(new Tuple<string, int, int>(ssn, row, row + 2));
                //SubsectionAddress.Add(ssn, row);
                ws.Cells[row, col].Value = ssn;
                ws.Cells[row, col, row, col + 2].Merge = true;
                ws.Cells[row, col, row, col + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                row += 2;
            }
            //TODO Создать метод возвращающий первую и последнюю строку в подразделе
            //TODO Объединить 3 колонки в одну для названия подраздела
            //TODO както распределять позиции по подразделам (тоже метод написать)
        }
        private Tuple<string, int, int> GetSubsectionTuple(string SubsectionName)
        {
            return SubsectionsAddress.Where(ss => ss.Item1.Equals(SubsectionName)).FirstOrDefault();
        }
        private void UpdateSubsectionAddress(string SubsectionName, Tuple<string, int, int> NewSubsection)
        {
            List<Tuple<string, int, int>> SSA = new List<Tuple<string, int, int>>();
            foreach (var ssa in SubsectionsAddress)
            {
                if (!ssa.Item1.Equals(SubsectionName))
                {
                    SSA.Add(NewSubsection);
                }
                SSA.Add(ssa);
            }
            SubsectionsAddress.Clear();
            SubsectionsAddress = SSA;
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
                    //var test = GetLastNotNullCol(ws) + 1;
                    //var test1 = ws.Cells[1, GetLastNotNullCol(ws) + 1].Value;
                    ws.Cells[1, GetLastNotNullCol(ws) + 1].Value = spec.DrawingNumber;//Добавляем в столбец номер очередного чертежа
                }
                ///Цикл по подразделам
                foreach (var ss in spec.Positions.Select(p => p.Subsection).Distinct())
                {
                    ///Цикл по разделам
                    foreach (var p in spec.Positions.Where(p => p.Section == ws.Name && p.Subsection == ss))
                    {
                        //Проверка на наличие записей впринципе
                        if (ws.Cells["A2"].Value == null)
                        {
                            //Не нашли записей, добавляем все подряд
                            ws.Cells[row, 1].Value = p.Subsection;//Добавляем подраздел
                            SubsectionsAddress.Add(new Tuple<string, int, int>(p.Subsection, row, row + 1));
                            row++;
                            ws.Cells[row, 2].Value = p.Name;//Добавляем название
                            ws.Cells[row, col].Value = p.Quantity;//Добавляем количество
                            row++;
                        }
                        else
                        {
                            string existSubsectionRow = ExistName(ws.Cells[$"A2:A{ws.Dimension.End.Row}"], ss);
                            if (string.IsNullOrEmpty(existSubsectionRow))
                            {
                                //Не нашли подраздела, добавляем в конец
                                row = GetLastNotNullRow(ws) + 1;
                                ws.SetValue(row, 1, p.Subsection);
                                ws.SetValue(row + 1, 2, p.Name);
                                SubsectionsAddress.Add(new Tuple<string, int, int>(p.Subsection, row, row + 1));
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
                                //Нашли подраздел
                                //TODO заполнять имеющимися на листе подразделами, т.к. они могут быть уже добавлены в предыдущих чертежах а в этом еще нет
                                if (GetSubsectionTuple(ss) == null)
                                {
                                    int sr = Convert.ToInt32(new string(existSubsectionRow.ToCharArray().Where(r => char.IsDigit(r)).ToArray()));
                                    int er = GetLastNotNullRowInSubsection(ws, sr);
                                    if (sr == er)
                                    {
                                        er = GetLastNotNullRow(ws);
                                    }
                                    SubsectionsAddress.Add(new Tuple<string, int, int>(ss, sr, er));
                                }

                                string existRow = ExistName(ws.Cells[$"B{GetSubsectionTuple(ss).Item2}:B{GetSubsectionTuple(ss).Item3}"], p.Name);
                                //Делаем поиск на существующее название
                                if (string.IsNullOrEmpty(existRow))
                                {
                                    //Не нашли одинакового, добавляем в конец
                                    var SsTuple = GetSubsectionTuple(ss);
                                    OffsetDownRange(ws, SsTuple.Item3 + 1);//Смещаем нижние подразделы ниже(если есть)
                                    row = SsTuple.Item3 + 1;//GetLastNotNullRow(ws) + 1;
                                    ws.SetValue(row, 2, p.Name);

                                    UpdateSubsectionAddress(ss, new Tuple<string, int, int>(ss, SsTuple.Item2, SsTuple.Item3 + 1));
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
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, ex.Message);
                Console.WriteLine("Возникла ошибка. Подробности см. в логе");
            }
        }

        private void OffsetDownRange(ExcelWorksheet ws, int row)
        {
            bool existvalues = false;
            for (int r = row; r < ws.Dimension.End.Row; r++)
            {
                for (int c = 1; c < ws.Dimension.End.Column; c++)
                {
                    if (ws.GetValue(r, c) != null)
                    {
                        existvalues = true;
                        break;
                    }
                }
            }
            if (existvalues)
            {
                var SsClone = SubsectionsAddress.Where(ssa => ssa.Item2 >= row).ToList();
                ws.Cells[row, ws.Dimension.Start.Column, ws.Dimension.End.Row, ws.Dimension.End.Column].Offset(row + 1, 1);//Проверить смещает ли
                //НЕ смещает!!
                foreach (var ssa in SsClone)
                {
                    UpdateSubsectionAddress(ssa.Item1, new Tuple<string, int, int>(ssa.Item1, ssa.Item2 + 1, ssa.Item3 + 1));
                }
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
                var t1 = ws.GetValue(row, 2);
                var t2 = ws.GetValue(row - 1, 2);
                var t3 = ws.GetValue(row, 1);
                if (ws.GetValue(row, 2) == null && ws.GetValue(row - 1, 2) != null && ws.GetValue(row, 1) == null)
                {
                    row_ = row - 1;
                    break;
                }
            }
            return row_;
        }
        private int GetLastNotNullRowInSubsection(ExcelWorksheet ws, int startrow)
        {
            int endrow = startrow;
            for (int r = startrow + 1; r < ws.Dimension.End.Row; r++)
            {
                if (ws.GetValue(r, 1) != null)
                {
                    endrow = r - 1;
                    break;
                }
            }
            return endrow;
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
