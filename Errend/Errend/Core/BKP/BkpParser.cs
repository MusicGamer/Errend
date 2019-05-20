using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AngleSharp.Dom.Html;

namespace Errend.Core.BKP
{
    class BkpParser : IParser<string[]>
    {
        public string[] Parse(IHtmlDocument document)
        {
            var list = new List<string>();
            var items = document.QuerySelectorAll("td").Where(item => item.ClassName != null && (item.ClassName.Contains("views-field views-field-field-shedule-vessel views-align-center") || item.ClassName.Contains("views-field views-field-field-shedule-outbound-voyage views-align-center")));
            foreach (var item in items)
            {
                list.Add(item.TextContent);
            }

            return list.ToArray();
        }
    }
}
