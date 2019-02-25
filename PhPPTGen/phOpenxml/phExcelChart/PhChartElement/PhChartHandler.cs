using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhChartHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.Chart chart = new C.Chart();
			chart.Append(new C.AutoTitleDeleted() { Val = Boolean.Parse((string)format["autoTitleDeleted"]) });
			chart.Append(new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap });
			chart.Append(new C.ShowDataLabelsOverMaximum() { Val = false });
			return chart;
		}
	}
}
