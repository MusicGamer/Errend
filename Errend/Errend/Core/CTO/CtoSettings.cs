
namespace Errend.Core.CTO
{
    class CtoSettings : IParserSettings
    {
        public string BaseUrl { get; set; } = "http://cto.od.ua/ru/rep/a.pub/vessel_call.html";
        public string Prefix { get; set; } = "{CurrentId}";
    }
}
