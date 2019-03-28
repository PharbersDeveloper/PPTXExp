using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhChartShapePropertiesHasFill : PhBaseElementHandler {

		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.ChartShapeProperties chartShapeProperties = new C.ChartShapeProperties();

			A.SolidFill solidFill = SolidFillFuncMap[(string)format["solidFill"]["type"]](format["solidFill"]); 

			A.Outline outline = new A.Outline();

			outline.Append(SolidFillFuncMap[(string)format["outline"]["type"]](format["outline"]));
			A.EffectList effectList = new A.EffectList();

			chartShapeProperties.Append(solidFill);
			chartShapeProperties.Append(outline);
			chartShapeProperties.Append(effectList);
			return chartShapeProperties;
		}

		private delegate A.SolidFill GetSolidFill(JToken format);

		private readonly Dictionary<string, GetSolidFill> SolidFillFuncMap = new Dictionary<string, GetSolidFill>() {
			{"color", (format => new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() { Val = int.Parse((string)format["alpha"]) }) { Val = new HexBinaryValue((string)format["color"]) }) ) },
			{"Scheme", (format => new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }) ) }
			
		};
	}
}
