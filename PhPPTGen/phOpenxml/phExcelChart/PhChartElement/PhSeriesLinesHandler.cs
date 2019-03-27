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
	class PhSeriesLinesHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.SeriesLines seriesLines = new C.SeriesLines();
			return seriesLines;
		}
	}
}
