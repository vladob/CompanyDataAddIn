using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompanyDataAddIn
{
    public class TitleRowConfig
    {
        public string BackgroundColor { get; set; }
        public FontConfig Font { get; set; }
        public string LabelEn { get; set; }
        public string LabelSk { get; set; }
    }

    public class Config
    {
        public TitleRowConfig TitleRow { get; set; }
        public List<DataConfig> DataConfig { get; set; }
    }
    public class DataConfig
    {
        public int Order { get; set; }
        public string ApiKey { get; set; }
        public string DataFormatting { get; set; }
        public FontConfig Font { get; set; }
        public string BackgroundColor { get; set; }
        public string LabelEn { get; set; }
        public string LabelSk { get; set; }
        public string LookupDictionary { get; set; }
    }

    public class LegalForms
    {
        public int Code { get; set; }
        public string TitleEng { get; set; }
        public string TitleSk { get; set; }
    }

    public class OrganizationSizes
    {
        public int Code { get; set; }
        public string TitleEng { get; set; }
        public string TitleSk { get; set; }
    }

    public class OwnershipTypes
    {
        public int Code { get; set; }
        public string TitleEng { get; set; }
        public string TitleSk { get; set; }
    }
    public class SkNace
    {
        public int Code { get; set; }
        public string TitleEng { get; set; }
        public string TitleSk { get; set; }
    }

}
