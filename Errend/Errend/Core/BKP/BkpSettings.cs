

namespace Errend.Core.BKP
{
    class BkpSettings : IParserSettings
    {
        public string BaseUrl { get; set; } = "http://www.bkport.com/ru/shedule";
        public string Prefix { get; set; } = "{CurrentId}";
    }
}
