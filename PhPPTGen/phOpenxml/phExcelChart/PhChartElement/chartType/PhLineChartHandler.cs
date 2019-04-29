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
	class PhLineChartHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.LineChart lineChart = new C.LineChart();
			lineChart.Append(new C.Grouping() { Val = (C.GroupingValues)Enum.Parse(typeof(C.GroupingValues), (string)format["grouping"]) });
			lineChart.Append(new C.VaryColors() { Val = Boolean.Parse((string)format["varyColors"]) });
			foreach (List<string> values in content.Series) {
				lineChart.Append(CreateLineChartSeries((uint)content.SeriesForIndex.IndexOf(values), values, content, format));
			}

			lineChart.Append(new C.Smooth() { Val = Boolean.Parse((string)format["varyColors"]) });
			foreach (string id in (JArray)format["axisID"]) {
				lineChart.Append(new C.AxisId() { Val = (UInt32Value)uint.Parse(id) });
			}

			return lineChart;
		}

		private C.LineChartSeries CreateLineChartSeries(uint index, List<string> values, PhChartContent content, JToken format) {
			C.LineChartSeries lineChartSeries = new C.LineChartSeries();
			lineChartSeries.Append(new C.Index() { Val = (UInt32Value)index });
			lineChartSeries.Append(new C.Order() { Val = (UInt32Value)index });
			lineChartSeries.Append(CreateSeriesText(content.CategoryLabels[(int)index], "Sheet1!$A$" + (2 + index)));
			lineChartSeries.Append(AppendOneElement(content, ((JArray)format["seriesChartShapeProperties"])[(int)index]));
			lineChartSeries.Append(AppendOneElement(content, ((JArray)format["markets"])[(int)index]));
			//lineChartSeries.Append(new C.Marker(new C.Symbol() { Val = (C.MarkerStyleValues)Enum.Parse(typeof(C.MarkerStyleValues), (string)format["marker"]) }));
			//todo: Sheet1!$A$2:$A$6 还需要一致, 要可能会有ppt损坏的问题
			lineChartSeries.Append(new C.CategoryAxisData(CreateStringReference(content.SeriesLabels, "Sheet1!$A$2:$A$6")));
			lineChartSeries.Append(new C.Values(CreateNumberReference(values, "Sheet1!$B$2:$D$2", (string)format["numFormat"])));
			lineChartSeries.Append(new C.Smooth() { Val = Boolean.Parse((string)format["varyColors"]) });
			lineChartSeries.Append(CreateLineSerExtensionList());
			return lineChartSeries;
		}

		//private C.SeriesText CreateSeriesText(string value, string formulaValue) {
		//	C.SeriesText seriesText = new C.SeriesText();
		//	seriesText.Append(CreateStringReference(new List<string> { value }, formulaValue));
		//	return seriesText;
		//}

		//private C.StringReference CreateStringReference(List<string> values, string formulaValue) {
		//	C.StringReference stringReference = new C.StringReference();
		//	C.Formula formula = new C.Formula {
		//		//excel上的位置	
		//		Text = formulaValue // "Sheet1!$B$1:$D$1";
		//	};

		//	C.StringCache stringCache = new C.StringCache();
		//	stringCache.Append(new C.PointCount() { Val = new UInt32Value((uint)values.Count) });
		//	foreach (string value in values) {
		//		C.StringPoint stringPoint = new C.StringPoint() { Index = (UInt32Value)(uint)values.ToList().IndexOf(value) };
		//		C.NumericValue numericValue = new C.NumericValue {
		//			Text = value
		//		};
		//		stringPoint.Append(numericValue);
		//		stringCache.Append(stringPoint);
		//	}

		//	stringReference.Append(formula);
		//	stringReference.Append(stringCache);

		//	return stringReference;
		//}

		//private C.NumberReference CreateNumberReference(List<string> values, string formulaValue, string numFormat) {

		//	C.NumberReference numberReference = new C.NumberReference();
		//	C.Formula formula = new C.Formula {
		//		Text = formulaValue // "Sheet1!$B$2:$D$2"
		//	};

		//	C.NumberingCache numberingCache = new C.NumberingCache();

		//	numberingCache.Append(new C.FormatCode { Text = numFormat });
		//	numberingCache.Append(new C.PointCount() { Val = (UInt32Value)(uint)values.Count });
		//	foreach (string value in values) {
		//		C.NumericPoint numericPoint = new C.NumericPoint() { Index = (UInt32Value)(uint)values.ToList().IndexOf(value) };
		//		C.NumericValue numericValue = new C.NumericValue { Text = value };
		//		numericPoint.Append(numericValue);
		//		numberingCache.Append(numericPoint);
		//	}
		//	numberReference.Append(formula);
		//	numberReference.Append(numberingCache);

		//	return numberReference;
		//}

		private C.LineSerExtensionList CreateLineSerExtensionList() {
			C.LineSerExtensionList lineSerExtensionList = new C.LineSerExtensionList();

			C.LineSerExtension lineSerExtension1 = new C.LineSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			lineSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-AA4A-407C-84E1-01D5F8660615}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			lineSerExtension1.Append(openXmlUnknownElement3);

			lineSerExtensionList.Append(lineSerExtension1);
			return lineSerExtensionList;
		}
	}
}
