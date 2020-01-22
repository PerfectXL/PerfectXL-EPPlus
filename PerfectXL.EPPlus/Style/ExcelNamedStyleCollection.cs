using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    public class ExcelNamedStyleCollection : IEnumerable<ExcelNamedStyle>
    {
        private readonly List<ExcelNamedStyle> _list;
        private readonly ExcelStyles _styles;

        internal ExcelNamedStyleCollection(ExcelStyleCollection<ExcelNamedStyleXml> list, ExcelStyles styles)
        {
            _styles = styles;
            _list = list.Select(x => new ExcelNamedStyle(x)).ToList();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public IEnumerator<ExcelNamedStyle> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public ExcelNamedStyle Create(string name)
        {
            return _styles.CreateNamedStyle(name, null);
        }

        public ExcelNamedStyle Create(string name, ExcelStyle template)
        {
            return _styles.CreateNamedStyle(name, template);
        }
    }
}