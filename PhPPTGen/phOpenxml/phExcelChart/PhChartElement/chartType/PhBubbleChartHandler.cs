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
	class PhBubbleChartHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.BubbleChart bubbleChart = new C.BubbleChart();
			bubbleChart.Append(new C.VaryColors() { Val = Boolean.Parse((string)format["varyColors"]) });
			bubbleChart.Append(
				new C.BubbleScale() { Val = (UInt32Value)uint.Parse((string)format["scale"]) }, 
				new C.ShowNegativeBubbles() { Val = Boolean.Parse((string)format["showNegative"]) });
			foreach (List<string> values in content.Series) {
				bubbleChart.Append(CreateBubbleChartSeries((uint)content.Series.IndexOf(values), values, content, format));
			}

			foreach (string id in (JArray)format["axisID"]) {
				bubbleChart.Append(new C.AxisId() { Val = (UInt32Value)uint.Parse(id) });
			}
			return bubbleChart;
		}

		private C.BubbleChartSeries CreateBubbleChartSeries(uint index, List<string> values, PhChartContent content, JToken format) {
			C.BubbleChartSeries bubbleChartSeries = new C.BubbleChartSeries();
			bubbleChartSeries.Append(new C.Index() { Val = (UInt32Value)index });
			bubbleChartSeries.Append(new C.Order() { Val = (UInt32Value)index });
			bubbleChartSeries.Append(CreateSeriesText(content.CategoryLabels[(int)index], "Sheet1!$A$" + (2 + index)));
			bubbleChartSeries.Append(AppendOneElement(content, ((JArray)format["seriesChartShapeProperties"])[(int)index]));
			bubbleChartSeries.Append(new C.InvertIfNegative() { Val = false });
			bubbleChartSeries.Append(AppendOneElement(content, format["serisDataLables"]));
			bubbleChartSeries.Append(new C.XValues(CreateNumberReference(new List<string>() { values[0] }, "Sheet1!$B$2", (string)format["xNumFormat"])));
			bubbleChartSeries.Append(new C.YValues(CreateNumberReference(new List<string>() { values[1] }, "Sheet1!$C$2", (string)format["yNumFormat"])));
			bubbleChartSeries.Append(new C.BubbleSize(CreateNumberReference(new List<string>() { values[2] }, "Sheet1!$D$2", (string)format["sizeNumFormat"])));
			bubbleChartSeries.Append(new C.Bubble3D() { Val = false });
			bubbleChartSeries.Append(CreateBubbleSerExtensionList());
			return bubbleChartSeries;
		}

		private C.BubbleSerExtensionList CreateBubbleSerExtensionList() {
			C.BubbleSerExtensionList bubbleSerExtensionList = new C.BubbleSerExtensionList();

			C.BubbleSerExtension bubbleSerExtension = new C.BubbleSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			bubbleSerExtension.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-2E80-4EC6-AF35-498A48A9743D}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			bubbleSerExtension.Append(openXmlUnknownElement);

			bubbleSerExtensionList.Append(bubbleSerExtension);
			return bubbleSerExtensionList;
		}
	}
}
