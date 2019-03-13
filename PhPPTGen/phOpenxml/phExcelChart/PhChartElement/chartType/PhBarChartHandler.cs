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
	class PhBarChartHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.BarChart barChart = new C.BarChart();

			//这儿定义条形图还是柱状图，由format属性决定
			barChart.Append(new C.BarDirection() { Val = (C.BarDirectionValues)Enum.Parse(typeof(C.BarDirectionValues), (string)format["direction"]) });
			//图表二级分类，堆积图，百分比堆积图之类，由format属性决定
			barChart.Append(new C.BarGrouping() { Val = (C.BarGroupingValues)Enum.Parse(typeof(C.BarGroupingValues), (string)format["grouping"]) });
			barChart.Append(new C.VaryColors() { Val = Boolean.Parse((string)format["varyColors"]) });
			foreach (List<string> values in content.Series) {
				barChart.Append(CreateBarChartSeries((uint)content.Series.IndexOf(values), values, content, format));
			}

;
			//宽度,深度,重叠 由format属性决定
			barChart.Append(new C.GapWidth() { Val = (UInt16Value)uint.Parse((string)format["gapWidth"]) });
			barChart.Append(new C.GapDepth() { Val = (UInt16Value)uint.Parse((string)format["gapDepth"]) });
			barChart.Append(new C.Overlap() { Val = (SByteValue)int.Parse((string)format["overlap"]) });
			//轴id 由format属性决定,list
			foreach (string id in (JArray)format["axisID"]) {
				barChart.Append(new C.AxisId() { Val = (UInt32Value)uint.Parse(id) });
			}
			return barChart;
		}

		private C.BarChartSeries CreateBarChartSeries(uint index, List<string> values, PhChartContent content, JToken format) {
			C.BarChartSeries barChartSeries = new C.BarChartSeries();
			barChartSeries.Append(new C.Index() { Val = (UInt32Value)index });
			barChartSeries.Append(new C.Order() { Val = (UInt32Value)index });
			//todo: 不同图表formulaValue生成方式不同
			barChartSeries.Append(CreateSeriesText(content.CategoryLabels[(int)index], "Sheet1!$A$" + (2 + index)));
			barChartSeries.Append(AppendOneElement(content, ((JArray)format["seriesChartShapeProperties"])[(int)index]));

			//是否补色填充负数，以后可能需要可定义
			barChartSeries.Append(new C.InvertIfNegative() { Val = false });
			barChartSeries.Append(AppendOneElement(content, format["serisDataLables"]));	

			//todo: 不同图表formulaValue生成方式不同
			barChartSeries.Append(new C.CategoryAxisData(CreateStringReference(content.seriesLabels, "Sheet1!$B$1:$D$1")));
			barChartSeries.Append(new C.Values(CreateNumberReference(values, "Sheet1!$B$2:$D$2", (string)format["numFormat"])));

			barChartSeries.Append(CreateBarSerExtensionList());
			return barChartSeries;
		}


		private C.BarSerExtensionList CreateBarSerExtensionList() {
			C.BarSerExtensionList barSerExtensionList = new C.BarSerExtensionList();

			C.BarSerExtension barSerExtension = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			barSerExtension.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-AA4A-407C-84E1-01D5F8660615}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			barSerExtension.Append(openXmlUnknownElement);

			barSerExtensionList.Append(barSerExtension);
			return barSerExtensionList;
		}


	}
}
