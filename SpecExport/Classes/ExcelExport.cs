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
                        Console.WriteLine($"{NameDoc} сгенерирован!");
                        log.Trace($"{NameDoc} сгенерирован!");
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
    }
}
