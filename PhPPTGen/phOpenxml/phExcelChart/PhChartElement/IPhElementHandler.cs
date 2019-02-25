using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	interface IPhElementHandler {
		OpenXmlCompositeElement CreateElement(PhChartContent content, JToken format);
	}
}
