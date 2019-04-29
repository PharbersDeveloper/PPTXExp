using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement{
	class PhComboChartHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			var numberMap = new Dictionary<string, int>() {
				{"all", content.Series.Count()},
				{"default", 0}
			};
			var take = numberMap[(string)format["takeBase"]] + int.Parse((string)format["take"]);
			var skip = numberMap[(string)format["skipBase"]] + int.Parse((string)format["skip"]);
			
			var localContent = new PhChartContent(content.SeriesForIndex,
				content.Series.Skip(skip).Take(take).ToList<List<string>>(),
				content.CategoryLabels,
				content.SeriesLabels,
				content.DataLabels
				);
			content.Series.Skip(skip).Take(take);
			return GetHandler((string)format["chartType"]["factory"]).CreateElement(localContent, format["chartType"]);
		}
	}
}
