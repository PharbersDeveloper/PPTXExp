using PhPPTGen.phOpenxml.phExcelChart.DO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json.Linq;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhChartPartsHandler {
		public PhChartContent Content { get; set; } = new PhChartContent();
		public JToken Format { get; set; } = null;

		public void CreateChartPart(ChartPart chartPart) {
			C.ChartSpace chartSpace = new C.ChartSpace();
			CreatrChartSpace(chartSpace);
			chartPart.ChartSpace = chartSpace;
		}

		private void CreatrChartSpace(C.ChartSpace chartSpace) {
			chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			chartSpace.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
			C.Date1904 date19041 = new C.Date1904() { Val = false };
			C.EditingLanguage editingLanguage = new C.EditingLanguage() { Val = "zh-CN" };
			C.RoundedCorners roundedCorners = new C.RoundedCorners() { Val = false };

			AlternateContent alternateContent = new AlternateContent();
			alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

			AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
			alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
			C14.Style style1 = new C14.Style() { Val = 102 };

			alternateContentChoice1.Append(style1);

			AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
			C.Style style2 = new C.Style() { Val = 2 };

			alternateContentFallback1.Append(style2);

			alternateContent.Append(alternateContentChoice1);
			alternateContent.Append(alternateContentFallback1);
			C.Chart chart = CreateChart();
			
			C.ShapeProperties shapeProperties = new C.ShapeProperties();
			CreateShapeProperties(shapeProperties);
			C.TextProperties textProperties = new C.TextProperties();
			CreateTextProperties(textProperties);
			C.PrintSettings printSettings = new C.PrintSettings();
			C.HeaderFooter headerFooter1 = new C.HeaderFooter();
			C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
			C.PageSetup pageSetup2 = new C.PageSetup();

			printSettings.Append(headerFooter1);
			printSettings.Append(pageMargins2);
			printSettings.Append(pageSetup2);

			chartSpace.Append(date19041);
			chartSpace.Append(editingLanguage);
			chartSpace.Append(roundedCorners);
			chartSpace.Append(alternateContent);
			chartSpace.Append(chart);
			chartSpace.Append(shapeProperties);
			chartSpace.Append(textProperties);
			chartSpace.Append(printSettings);
		}

		private C.Chart CreateChart() {
			return (C.Chart)new PhChartHandler().CreateElement(Content, Format);
		}

		private C.ShapeProperties CreateShapeProperties(C.ShapeProperties shapeProperties) {
			A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex() { Val = new HexBinaryValue("FFFFFF") });

			A.Outline outline = new A.Outline();
			A.NoFill outlineNoFill = new A.NoFill();

			outline.Append(outlineNoFill);
			A.EffectList effectList = new A.EffectList();

			shapeProperties.Append(solidFill);
			shapeProperties.Append(outline);
			shapeProperties.Append(effectList);
			return shapeProperties;
		}

		private C.TextProperties CreateTextProperties(C.TextProperties textProperties) {
			A.BodyProperties bodyProperties = new A.BodyProperties();
			A.ListStyle listStyle = new A.ListStyle();

			A.Paragraph paragraph = new A.Paragraph();

			A.ParagraphProperties paragraphProperties = new A.ParagraphProperties();
			A.DefaultRunProperties defaultRunProperties = new A.DefaultRunProperties();

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
