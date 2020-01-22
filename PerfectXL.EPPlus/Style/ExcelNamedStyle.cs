using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    public class ExcelNamedStyle
    {
        private readonly ExcelNamedStyleXml _xml;

        internal ExcelNamedStyle(ExcelNamedStyleXml xml)
        {
            _xml = xml;
        }

        public int BuiltinId
        {
            get => _xml.BuiltinId;
            set => _xml.BuiltinId = value;
        }

        public bool CustomBuiltin
        {
            get => _xml.CustomBuiltin;
            set => _xml.CustomBuiltin = value;
        }

        public string Name
        {
            get => _xml.Name;
            set => _xml.Name = value;
        }

        public ExcelStyle Style
        {
            get => _xml.Style;
            set => _xml.Style = value;
        }
    }
}