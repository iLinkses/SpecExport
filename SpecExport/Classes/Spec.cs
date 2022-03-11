using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecExport.Classes
{
    /// <summary>
    /// Класс для стукруры спецификации
    /// </summary>
    public class Spec
    {
        private string fileName;
        /// <summary>
        /// Возвращает название файла без пути
        /// </summary>
        /// <remarks>Передавать сюда полный путь!</remarks>
        public string FileName
        {
            get { return fileName; }
            set
            {
                fileName = value.Remove(0, (value.Length - 1) - ((value.Length - 1) - value.LastIndexOf(@"\") - 1));
            }
        }

        /// <summary>
        /// Возвращает номер чертежа из названия чертежа
        /// </summary>
        /// <example>"30_НазваниеКакоготоЧертежа" вернет 30</example>
        public int? DrawingNumber
        {
            get
            {
                if (!string.IsNullOrEmpty(FileName))
                {
                    return Convert.ToInt32(FileName.Remove(FileName.IndexOf("_")));
                }
                else return null;
            }
        }

        public List<Detail> Positions { get; set; } = new List<Detail>();
        public class Detail
        {
            /// <summary>
            /// Для разделения на листы(книги)
            /// </summary>
            public string Section { get; set; }
            /// <summary>
            /// Для разделения на подразделы на листе
            /// </summary>
            public virtual string Subsection { get { return System.Text.RegularExpressions.Regex.Match(Name, @"[а-яА-Я]+").Value; } set { } }
            public string Format { get; set; }
            public string Zone { get; set; }
            public string Position { get; set; }
            public string Designation { get; set; }
            public string Name { get; set; }
            public decimal Quantity { get; set; }
            public string Note { get; set; }
        }
    }
    public class ExcelSpec
    {
        public List<ExcelDetail> Details { get; set; } = new List<ExcelDetail>();
        public class ExcelDetail : Spec.Detail
        {
            public string FileName { get; set; }
            public int? DrawingNumber { get; set; }
            private string subsection;
            public override string Subsection { get => base.Subsection; set => subsection = value; }
        }
    }
}
