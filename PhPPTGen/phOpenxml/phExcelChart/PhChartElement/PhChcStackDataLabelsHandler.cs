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
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhChcStackedDataLabelHandler : PhDataLabelsHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.DataLabels dataLabels = new C.DataLabels();
			var index = 0;
			foreach (string s in content.SeriesLabels) {
				dataLabels.Append(CreateDataLabel(content.DataLabels.First()[index], format, (uint)index));
				index ++;
			}
			content.DataLabels.RemoveAt(0);
			//标签属性，以后可以由format属性定义
			dataLabels.Append(new C.DataLabelPosition() { Val = (C.DataLabelPositionValues)Enum.Parse(typeof(C.DataLabelPositionValues), (string)format["position"]) });
			dataLabels.Append(new C.ShowLegendKey() { Val = false });
			dataLabels.Append(new C.ShowValue() { Val = Boolean.Parse((string)format["showValue"]) });
			dataLabels.Append(new C.ShowCategoryName() { Val = Boolean.Parse((string)format["showCategoryName"]) });
			dataLabels.Append(new C.ShowSeriesName() { Val = Boolean.Parse((string)format["showSeriesName"]) });
			dataLabels.Append(new C.ShowPercent() { Val = Boolean.Parse((string)format["showPercent"]) });
			dataLabels.Append(new C.ShowBubbleSize() { Val = false });
			dataLabels.Append(new C.ShowLeaderLines() { Val = false });

			dataLabels.Append(CreateDLblsExtensionList());
			return dataLabels;
		}

		protected virtual C.DataLabel CreateDataLabel(string content, JToken format, UInt32 index) {
			C.DataLabel dataLabel = new C.DataLabel(new C.Index() { Val = (UInt32Value)index });

			C.ChartText chartText = new C.ChartText();
			C.RichText richText = new C.RichText(new A.BodyProperties(), new A.ListStyle());
			A.Paragraph paragraph = new A.Paragraph();
			A.Run run = new A.Run(new A.RunProperties() { Language = "en-US", AlternativeLanguage = "zh-CN" });
			A.Text text = new A.Text {
				Text = content
			};
			run.Append(text);
			paragraph.Append(run);
			richText.Append(paragraph);
			chartText.Append(richText);

			dataLabel.Append(chartText);
			dataLabel.Append(new C.DataLabelPosition() { Val = (C.DataLabelPositionValues)Enum.Parse(typeof(C.DataLabelPositionValues), (string)format["position"]) });
			dataLabel.Append(new C.ShowLegendKey() { Val = false });
			dataLabel.Append(new C.ShowValue() { Val = true });
			dataLabel.Append(new C.ShowCategoryName() { Val = false });
			dataLabel.Append(new C.ShowSeriesName() { Val = false });
			dataLabel.Append(new C.ShowPercent() { Val = false });
			dataLabel.Append(new C.ShowBubbleSize() { Val = false });

			dataLabel.Append(CreateDLblExtensionList());
			return dataLabel;
		}
		protected virtual C.DLblExtensionList CreateDLblExtensionList() {
			C.DLblExtensionList dLblExtensionList = new C.DLblExtensionList();

			C.DLblExtension dLblExtension1 = new C.DLblExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
			dLblExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");

			C.DLblExtension dLblExtension2 = new C.DLblExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			dLblExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{0000000E-D872-438E-B420-8FE68279862A}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			dLblExtension2.Append(openXmlUnknownElement8);

			dLblExtensionList.Append(dLblExtension1);
			dLblExtensionList.Append(dLblExtension2);
			return dLblExtensionList;
		}
	}

}
