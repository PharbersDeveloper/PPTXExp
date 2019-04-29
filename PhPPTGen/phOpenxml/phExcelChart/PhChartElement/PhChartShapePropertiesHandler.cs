using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhChartShapePropertiesHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.ChartShapeProperties chartShapeProperties = new C.ChartShapeProperties();
			//自动填充，要指定填充用PhChartShapePropertiesHasFill
			//A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() { Val = int.Parse((string)format["alpha"]) }) { Val = new HexBinaryValue((string)format["solidFill"]) });

			A.Outline outline = new A.Outline();

			outline.Append(new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() { Val = int.Parse((string)format["alpha"]) }) { Val = new HexBinaryValue((string)format["outline"]) }));
			A.EffectList effectList = new A.EffectList();

			//chartShapeProperties.Append(solidFill);
			chartShapeProperties.Append(outline);
			chartShapeProperties.Append(effectList);
			return chartShapeProperties;


		}
	}
}
