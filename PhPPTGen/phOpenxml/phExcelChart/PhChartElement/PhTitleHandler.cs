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
	class PhTitleHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.Title title = new C.Title();
			string titleText = content.titleMap[(string)format["titleType"]];
			AppendChartText(title, titleText);
			C.Overlay overlay = new C.Overlay() { Val = false };
			title.Append(overlay);
			return title;
		}

		private void AppendChartText(C.Title title, string value) {
			C.ChartText chartText = new C.ChartText();

			C.RichText richText = new C.RichText();
			A.BodyProperties bodyProperties = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
			A.ListStyle listStyle = new A.ListStyle();

			A.Paragraph paragraph = new A.Paragraph();

			A.ParagraphProperties paragraphProperties = new A.ParagraphProperties();

			A.DefaultRunProperties defaultRunProperties = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

			A.SolidFill solidFill = new A.SolidFill();

			A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			A.LuminanceModulation luminanceModulation = new A.LuminanceModulation() { Val = 65000 };
			A.LuminanceOffset luminanceOffset = new A.LuminanceOffset() { Val = 35000 };

			schemeColor.Append(luminanceModulation);
			schemeColor.Append(luminanceOffset);

			solidFill.Append(schemeColor);
			A.LatinFont latinFont = new A.LatinFont() { Typeface = "+mn-lt" };
			A.EastAsianFont eastAsianFont = new A.EastAsianFont() { Typeface = "+mn-ea" };
			A.ComplexScriptFont complexScriptFont = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

			defaultRunProperties.Append(solidFill);
			defaultRunProperties.Append(latinFont);
			defaultRunProperties.Append(eastAsianFont);
			defaultRunProperties.Append(complexScriptFont);

			paragraphProperties.Append(defaultRunProperties);

			A.Run run = new A.Run();
			A.RunProperties runProperties = new A.RunProperties() { Language = "en-US", AlternativeLanguage = "zh-CN" };
			A.Text text = new A.Text {
				Text = value //似乎只需要更改这儿
			};

			run.Append(runProperties);
			run.Append(text);

			paragraph.Append(paragraphProperties);
			paragraph.Append(run);

			richText.Append(bodyProperties);
			richText.Append(listStyle);
			richText.Append(paragraph);

			chartText.Append(richText);
			title.Append(chartText);
		}
	}
}
