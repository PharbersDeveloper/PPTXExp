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
	class PhLegendHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.Legend legend = new C.Legend();
			legend.Append(new C.LegendPosition() { Val = (C.LegendPositionValues)Enum.Parse(typeof(C.LegendPositionValues), (string)format["legendPosition"]) });
			legend.Append(new C.Overlay() { Val = false });
			return legend;
		}
	}
}
