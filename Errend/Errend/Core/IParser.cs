using AngleSharp.Dom.Html;

namespace Errend.Core
{
    interface IParser<T> where T : class
    {
        T Parse(IHtmlDocument document);
    }
}
