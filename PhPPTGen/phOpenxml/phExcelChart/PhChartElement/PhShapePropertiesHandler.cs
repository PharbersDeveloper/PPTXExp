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
	class PhShapePropertiesHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.ShapeProperties shapeProperties = new C.ShapeProperties();

			A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() {
				Val = int.Parse((string)format["alpha"])}) { Val = new HexBinaryValue((string)format["solidFill"]) });

			A.Outline outline = new A.Outline();

			outline.Append(new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() {
				Val = int.Parse((string)format["alpha"])}) { Val = new HexBinaryValue((string)format["outline"]) }));
			A.EffectList effectList = new A.EffectList();

			shapeProperties.Append(solidFill);
			shapeProperties.Append(outline);
			shapeProperties.Append(effectList);
			return shapeProperties;
		}
	}
}
