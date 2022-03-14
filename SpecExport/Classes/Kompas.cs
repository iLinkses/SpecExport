using System;
using System.Collections.Generic;
using Kompas6API5;
using KAPITypes;
using Kompas6Constants;
using System.Data;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;
using NLog;

namespace SpecExport.Classes
{
    class Kompas
    {
        private NLog.Logger log = Program.log;
        private string DrawingsDirectory { get { return Properties.Settings.Default.DrawingsDirectory; } }
        private KompasObject kompas { get; set; } = null;
        private ksDocument2D doc2D { get; set; } = null;
        private ksSpecification iSpc { get; set; } = null;
        private List<Spec> Specs { get; set; } = new List<Spec>();
        private Spec Spec { get; set; }
        private string FullFileName { get; set; }
        private readonly Dictionary<int, string> SpecialCharacters;
        //{
        //    { 2, "\u00D8" }//Ø
        //};
        public Kompas()
        {
            log = LogManager.GetCurrentClassLogger();
            SpecialCharacters = FillDictionary();//Заполняем словарь символов
        }
        private Dictionary<int,string> FillDictionary()
        {
            Dictionary<int, string> Dict = new Dictionary<int, string>();
            try
            {
                foreach (var scs in Properties.Settings.Default.SpecialCharacters.Split(';'))
                {
                    var matches = new Regex(@"(?<Key>\d),(?<Value>.*)").Matches(scs);
                    foreach (Match sc in matches)
                    {
                        Dict.Add(int.Parse(sc.Groups["Key"].Value.Trim()), sc.Groups["Value"].Value.Trim());
                    }
                }
                log.Trace("Загрузили словарь спец. символов");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Неверная структура спец. символов, подробнее см. лог");
                log.Error(ex, ex.Message);
            }
            return Dict;
        }
        public void ExportSpec()
        {
            List<string> FileNames = GetFileInCatalog();
            if (FileNames.Count > 0)
            {
                LoadKompas();
                log.Trace("Проходимся по чертежам");
                foreach (var fn in FileNames)
                {
                    this.FullFileName = $@"{Directory.GetCurrentDirectory()}\{DrawingsDirectory}\{fn}";
                    //Открываем чертеж
                    OpenFile();
                    if (doc2D != null)
                    {
                        iSpc = (ksSpecification)doc2D.GetSpecification();
                        GetSpec(iSpc);
                        //Загрузили спецификацию, закрываем чертеж
                        CloseFile();
                    }
                }
                CloseKompas();
                if (Specs.Count > 0)
                {
                    log.Trace("Генерируем отчет");
                    ExcelExport excelExport = new ExcelExport(Specs);
                    excelExport.NewSaveExcel();
                }
            }
        }
        /// <summary>
        /// Получает названия файлов с чертежами в каталоге
        /// </summary>
        private List<string> GetFileInCatalog()
        {
            List<string> FileNames = new List<string>();
            if (!Directory.Exists($@"{Directory.GetCurrentDirectory()}\{DrawingsDirectory}"))
            {
                Console.WriteLine($"Не найдет каталог {DrawingsDirectory} проверьте имя каталога в конфиге");
                log.Error($"Не найдет каталог {DrawingsDirectory}");
                return null;
            }
            //var t = Directory.Exists(DrawingsDirectory);
            //var t1 = File.Exists($@"{DrawingsDirectory}\Чертеж — копия (2).cdw");
            //Directory.GetFiles(DrawingsDirectory);
            //string DrawingsDirectory = Properties.Settings.Default.DrawingsDirectory;
            foreach (var f in Directory.GetFiles($@"{Directory.GetCurrentDirectory()}\{DrawingsDirectory}", "*.cdw"))
            {
                FileNames.Add(Path.GetFileName(f));
                //Console.WriteLine(Path.GetFileName(f));
            }

            List<object[]> cr = new List<object[]>();
            foreach (var fn_ in FileNames) cr.Add(new object[] { fn_ });
            Program.WConsoleTable(new string[] { "Название чертежа" }, cr);

            string fn = string.Empty;
            FileNames.ForEach(new Action<string>(name => fn += "\n\t" + name));

            if (FileNames.Count > 0) log.Trace($"Список чертежей получен: {fn}");
            else log.Error($"Пустой каталог {DrawingsDirectory}");
            return FileNames;
        }
        private void LoadKompas()
        {
            if (kompas == null)
            {
#if __LIGHT_VERSION__
				Type t = Type.GetTypeFromProgID("KOMPASLT.Application.5");
#else
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
#endif
                kompas = (KompasObject)Activator.CreateInstance(t);

                log.Trace($"Открыли компас");
            }

            if (kompas != null)
            {
                kompas.Visible = true;
                kompas.ActivateControllerAPI();
            }
        }
        private void CloseKompas()
        {
            //Закрывает компас
            if (kompas != null)
            {
                kompas.Quit();
                Marshal.ReleaseComObject(kompas);
                log.Trace($"Закрыли компас");
            }
        }
        private void OpenFile()
        {
            var t = File.Exists(FullFileName);
            int type = kompas.ksGetDocumentTypeByName(FullFileName);
            switch (type)
            {
                case (int)DocType.lt_DocSheetStandart:  //2d документы
                case (int)DocType.lt_DocFragment:
                    doc2D = (ksDocument2D)kompas.Document2D();
                    if (doc2D != null)
                    {
                        doc2D.ksOpenDocument(FullFileName, false);
                        log.Trace($"Открыли чертеж {FullFileName}");
                    }
                    else
                    {
                        log.Error("Не удалось открыть чертеж");
                    }
                    break;
            }
        }
        private void CloseFile()
        {
            if (doc2D != null)
            {
                var t = doc2D.ksCloseDocument();
                log.Trace($"Закрыли чертеж {FullFileName}");
            }
        }
        private void GetSpec(ksSpecification specification)
        {
            DataTable dt = new DataTable();
            List<string> list = new List<string>();
            Spec = new Spec();
            Spec.FileName = FullFileName;
            //ksDocument2D doc = (ksDocument2D)kompas.Document2D();
            //ksSpcDocument spc = (ksSpcDocument)kompas.SpcActiveDocument();
            //if (doc != null && spc != null && spc.reference != 0)
            //{
            //ksSpecification specification = (ksSpecification)spc.GetSpecification();

            //см справку на ksCreateSpcIterator
            //spcObjType
            //- тип объектов:
            //0 - базовые,
            //1 - вспомогательные,
            //2 - базовые и вспомогательные из сортирован­ного массива,
            //3 - все объекты.
            ksIterator iter = (ksIterator)kompas.GetIterator();
            iter.ksCreateSpcIterator(null, 0, 0);
            if (iter.reference != 0 && specification != null)
            {
                int obj = iter.ksMoveIterator("F");
                if (obj != 0)
                {
                    List<string> ls = new List<string>();
                    do
                    {
                        //узнаем количество колонок у базового объекта спецификации
                        int count = specification.ksGetSpcTableColumn(null, 0, 0);

                        string buf = string.Format("Кол-во колонок = {0}", count);
                        ls.Clear();
                        //kompas.ksMessage(buf);
                        //Console.WriteLine(buf);
                        // пройдем по всем колонкам
                        for (int i = 1; i <= count; i++)
                        {
                            // для текущего номера определим тип колонки, номер исполнения и блок
                            ksSpcColumnParam spcColPar = (ksSpcColumnParam)kompas.GetParamStruct((short)StructType2DEnum.ko_SpcColumnParam);
                            if (specification.ksGetSpcColumnType(obj,   //объект спецификации
                                i,                                      // номер колонки, начиная с 1
                                spcColPar) == 1)
                            {
                                // возьмем текст
                                int columnType = spcColPar.columnType;
                                int ispoln = spcColPar.ispoln;
                                int blok = spcColPar.block;
                                buf = specification.ksGetSpcObjectColumnText(obj, columnType, ispoln, blok);
                                //kompas.ksMessage(buf);
                                //Console.WriteLine(buf);
                                ls.Add(buf);
                                // по типу колонки, номеру исполнения и блоку определим номер колонки
                                //int colNumb = specification.ksGetSpcColumnNumb(obj, //объект спецификации
                                //    spcColPar.columnType, spcColPar.ispoln, spcColPar.block);
                                //buf = string.Format("i = {0} colNumb = {1}", i, colNumb);
                                //kompas.ksMessage(buf);
                                //Console.WriteLine(buf);
                            }
                        }
                        var Detail = new Spec.Detail();
                        if (ls.Count > 0)
                        {
                            ///Отбрасываем "информационные" позиции без количества и примечания
                            if (!string.IsNullOrEmpty(ls[5]) || !string.IsNullOrEmpty(ls[6]))
                            {
                                Detail.Section = specification.ksGetSpcSectionName(obj);
                                Detail.Format = ls[0];
                                Detail.Zone = ls[1];
                                Detail.Position = ls[2];
                                Detail.Designation = ls[3];
                                Detail.Name = ReplaseKompasSymbols(ls[4]);//ls[4];
                                Detail.Quantity = !string.IsNullOrEmpty(ls[5]) ? Convert.ToDecimal(ls[5]) : Convert.ToDecimal(Regex.Match(ls[6], @"\d+[.,]?\d*").Value);///вытаскиваем количество из примечания если столбец количество пустой
                                Detail.Note = ls[6];
                            }
                        }
                        Spec.Positions.Add(Detail);
                    }
                    while ((obj = iter.ksMoveIterator("N")) != 0);
                    Specs.Add(Spec);

                    log.Trace($"Получили спецификацию по чертежу {FullFileName}");
                }
            }
            //}
            //else
            //    kompas.ksError("Спецификация должна быть текущей");
        }
        /// <summary>
        /// Заменяет вставки спецсимволов компаса на юникод/ascii
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        /// TODO сначала заполнить словарь <see cref="SpecialCharacters"/>
        private string ReplaseKompasSymbols(string text)
        {
            string newtext = text;
            try
            {
                //TODO там могут быть другие конструкции, которые тоже стоит по идее учитывать
                var matches = new Regex(@"(?<insert>@(?<Modifier>\*|\+)?(?<CharacterCode>\w*?)~)").Matches(text);
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        newtext = text.Replace(match.Groups["insert"].Value, SpecialCharacters[Convert.ToInt32(match.Groups["CharacterCode"].Value)]);//TODO посмотреть в документации "Таблица спецзнаков находится в приложении V основной справки КОМПАСа"
                                                                                                                                                      //Console.WriteLine(match.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, $"{ex.Message}\n\t{text}");
            }
            return newtext;
        }

        [Obsolete("Всякий мусор")]
        private void Bucket()
        {
            #region получается изменять штамп
            //ksStamp stamp = (ksStamp)doc2D.GetStamp();
            //string test = string.Empty;
            //if (stamp != null)
            //{
            //    if (stamp.ksOpenStamp() == 1)
            //    {
            //        stamp.ksColumnNumber(2);

            //        ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
            //        if (itemParam != null)
            //        {
            //            itemParam.Init();
            //            ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
            //            if (itemFont != null)
            //            {
            //                itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
            //                test = itemParam.s;
            //                //itemParam.s = "1111111";
            //                //doc2D.ksTextLine(itemParam);
            //            }
            //        }

            //        stamp.ksCloseStamp();
            //    }
            //}
            #endregion

            #region Получилось получить данные штампа
            //List<string> ls = new List<string>();

            //int[] myArr = new int[10]; // Коды строк чертежа
            //myArr[0] = 2; // Обозначение
            //myArr[1] = 1; // Наименование
            //myArr[2] = 5; // Масса
            //myArr[3] = 3; // Материал
            //myArr[4] = 110; // Разработал
            //myArr[5] = 111; // Проверил
            //myArr[6] = 113; // нач КБ
            //myArr[7] = 114; // Н контроль
            //myArr[8] = 115; // Утв
            //myArr[9] = 25; // Первичное прим.

            ////ksDocument2D Doc = (ksDocument2D)kompas.Document2D();
            ////Doc.ksOpenDocument((treeView1.SelectedNode.FullPath + @"\" + checkedListBox1.Text), true);
            //ksDocumentParam DocPm = (ksDocumentParam)kompas.GetParamStruct(35);
            //ksStamp st = (ksStamp)doc2D.GetStamp();
            //st.ksOpenStamp();
            //foreach (int n in myArr)
            //{
            //    st.ksColumnNumber(n);
            //    ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
            //    ksTextLineParam TextLine = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
            //    ksDynamicArray f = (ksDynamicArray)st.ksGetStampColumnText(n);
            //    if (f != null)
            //    {
            //        string str_stamp = "";
            //        int rr = f.ksGetArrayCount();// определяет количество строк
            //        for (int i1 = 0; i1 < f.ksGetArrayCount(); i1++)
            //        {
            //            f.ksGetArrayItem(i1, TextLine); // читает определенную строку строку
            //            ksDynamicArray f1 = (ksDynamicArray)TextLine.GetTextItemArr();
            //            f1.ksGetArrayItem(0, itemParam);

            //            str_stamp = str_stamp + itemParam.s;
            //        }
            //        ls.Add(str_stamp);
            //    }
            //    else
            //    {
            //        ls.Add("");
            //    }

            //} 
            #endregion

            //    if (kompaskompas != null)
            //    {
            //        openFileDialog.Filter = "Чертежи(*.cdw)|*.cdw|Фрагменты(*.frw)|*.frw|Модели(*.m3d)|*.m3d|Сборки(*.a3d)|*.a3d|Спецификации(*.spw)|*.spw";
            //        //if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            //        //{
            //        // Открыть документ с диска
            //        // первый параметр - имя открываемого файла
            //        // второй параметр указывает на необходимость выдачи запроса "Файл изменен. Сохранять?" при закрытии файла
            //        // третий параметр - указатель на IDispatch, по которому График вызывает уведомления об изменении своего состояния
            //        // ф-ия возвращает HANDLE открытого документа

            //        int type = kompas.ksGetDocumentTypeByName(openFileDialog.FileName);
            //        ksDocument3D doc3D;
            //        ksDocument2D doc2D;
            //        ksSpcDocument docSpc;
            //        ksDocumentTxt docTxt;
            //        switch (type)
            //        {
            //            case (int)DocType.lt_DocPart3D:         //3d документы
            //            case (int)DocType.lt_DocAssemble3D:
            //                doc3D = (ksDocument3D)kompas.Document3D();
            //                if (doc3D != null)
            //                    doc3D.Open(openFileDialog.FileName, false);
            //                break;
            //            case (int)DocType.lt_DocSheetStandart:  //2d документы
            //            case (int)DocType.lt_DocFragment:
            //                doc2D = (ksDocument2D)kompas.Document2D();
            //                if (doc2D != null)
            //                    doc2D.ksOpenDocument(openFileDialog.FileName, false);
            //                break;
            //            case (int)DocType.lt_DocSpc:                //спецификации
            //                docSpc = (ksSpcDocument)kompas.SpcDocument();
            //                if (docSpc != null)
            //                    docSpc.ksOpenDocument(openFileDialog.FileName, 0);
            //                break;
            //            case (int)DocType.lt_DocTxtStandart:        //текстовые документы
            //                docTxt = (ksDocumentTxt)kompas.DocumentTxt();
            //                if (docTxt != null)
            //                    docTxt.ksOpenDocument(openFileDialog.FileName, 0);
            //                break;
            //        }
            //        int err = kompas.ksReturnResult();
            //        if (err != 0)
            //            kompas.ksResultNULL();
            //        //}
            //    }
            //    else
            //    {
            //        Console.WriteLine("Объект не захвачен");
            //    }
        }
    }
}
