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
	class PhPieChartHandler : PhChartTypeBaseHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.PieChart pieChart = new C.PieChart();
			pieChart.Append(new C.VaryColors() { Val = Boolean.Parse((string)format["varyColors"]) });
			foreach (List<string> values in content.Series) {
				pieChart.Append(CreatePieChartseries((uint)content.Series.IndexOf(values), values, content, format));
			}
			pieChart.Append(new C.FirstSliceAngle() { Val = UInt16.Parse((string)format["angle"]) });

			return pieChart;
		}

		private C.PieChartSeries CreatePieChartseries(uint index, List<string> values, PhChartContent content, JToken format) {
			C.PieChartSeries pieChartSeries = new C.PieChartSeries();
			pieChartSeries.Append(new C.Index() { Val = (UInt32Value)index });
			pieChartSeries.Append(new C.Order() { Val = (UInt32Value)index });
			pieChartSeries.Append(CreateSeriesText(content.CategoryLabels[(int)index], "Sheet1!$A$" + (2 + index)));
			pieChartSeries.Append(AppendOneElement(content, ((JArray)format["seriesChartShapeProperties"])[(int)index]));
			pieChartSeries.Append(new C.CategoryAxisData(CreateStringReference(content.seriesLabels, "Sheet1!$B$1:$D$1")));
			pieChartSeries.Append(new C.Values(CreateNumberReference(values, "Sheet1!$B$2:$D$2", (string)format["numFormat"])));
			pieChartSeries.Append(CreateLineSerExtensionList());
			return pieChartSeries;
		}

		private C.PieSerExtensionList CreateLineSerExtensionList() {
			C.PieSerExtensionList pieSerExtensionList = new C.PieSerExtensionList();

			C.PieSerExtension pieSerExtension1 = new C.PieSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			pieSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-612B-4BED-92C2-B85AD71039D4}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			pieSerExtension1.Append(openXmlUnknownElement4);

			pieSerExtensionList.Append(pieSerExtension1);
			return pieSerExtensionList;
		}

		private C.DataPoint CreatedataPoint(uint index, JToken format) {
			C.DataPoint dataPoint = new C.DataPoint();
			dataPoint.Append(new C.Index() { Val = (UInt32Value)index }, new C.Bubble3D() { Val = false });
			dataPoint.Append(AppendOneElement(new PhChartContent(), ((JArray)format["pointChartShapeProperties"])[(int)index]));
			C.ExtensionList extensionList = new C.ExtensionList();
			C.Extension extension = new C.Extension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			extension.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");
			OpenXmlUnknownElement openXmlUnknownElement = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000004-612B-4BED-92C2-B85AD71039D4}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");
			extension.Append(openXmlUnknownElement);
			extensionList.Append(extension);
			dataPoint.Append(extensionList);
			return dataPoint;
		}
	}
}
