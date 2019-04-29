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
	class PhMarkerHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			return new C.Marker(
				new C.Symbol() { Val = (C.MarkerStyleValues)Enum.Parse(typeof(C.MarkerStyleValues), (string)format["symbol"]) },
				new C.Size() { Val = Byte.Parse((string)format["size"]) }
			);
		}
	}
}
