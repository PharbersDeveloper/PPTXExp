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
	class PhTextPropertiesHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.TextProperties textProperties = new C.TextProperties();
			A.BodyProperties bodyProperties = new A.BodyProperties();
			A.ListStyle listStyle = new A.ListStyle();

			A.Paragraph paragraph = new A.Paragraph();

			A.ParagraphProperties paragraphProperties = new A.ParagraphProperties();
			A.DefaultRunProperties defaultRunProperties = new A.DefaultRunProperties() {
				FontSize = int.Parse((string)format["fontSize"]) * 100,
				Bold = Boolean.Parse((string)format["bold"]),
				Italic = false,
				Underline = A.TextUnderlineValues.None,
				Strike = A.TextStrikeValues.NoStrike,
				Kerning = int.Parse((string)format["kerning"]) * 100,
				Baseline = 0
			};

			paragraphProperties.Append(defaultRunProperties);
			A.EndParagraphRunProperties endParagraphRunProperties = new A.EndParagraphRunProperties() { Language = "zh-CN" };

			paragraph.Append(paragraphProperties);
			paragraph.Append(endParagraphRunProperties);

			textProperties.Append(bodyProperties);
			textProperties.Append(listStyle);
			textProperties.Append(paragraph);
			return textProperties;
		}
	}
}
