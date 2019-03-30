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
	class PhDataLabelsForParasHandler : PhDataLabelsHandler {

		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format, params object[] paras) {
			Dictionary<string, GetDataLabel> GetDataLabelFuncMap = new Dictionary<string, GetDataLabel>() {
				{"chcBubble",  GetChcBubbleDataLabel}
			};
			C.DataLabels dataLabels = new C.DataLabels();
			dataLabels.Append(GetDataLabelFuncMap[(string)format["type"]](content, format, paras));

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

			dataLabel.Append(AppendOneElement(new PhChartContent(), format["layout"]));
			dataLabel.Append(new C.DataLabelPosition() { Val = (C.DataLabelPositionValues)Enum.Parse(typeof(C.DataLabelPositionValues), (string)format["position"]) });
			dataLabel.Append(new C.ShowLegendKey() { Val = false });
			dataLabel.Append(new C.ShowValue() { Val = Boolean.Parse((string)format["showValue"]) });
			dataLabel.Append(new C.ShowCategoryName() { Val = Boolean.Parse((string)format["showCategoryName"]) });
			dataLabel.Append(new C.ShowSeriesName() { Val = Boolean.Parse((string)format["showSeriesName"]) });
			dataLabel.Append(new C.ShowPercent() { Val = Boolean.Parse((string)format["showPercent"]) });
			dataLabel.Append(new C.ShowBubbleSize() { Val = false });

			dataLabel.Append(CreateDLblExtensionList());
			return dataLabel;
		}


		protected C.DLblExtensionList CreateDLblExtensionList() {
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

		private C.DataLabel GetChcBubbleDataLabel(PhChartContent content, JToken format, params object[] paras) {
			double size = double.Parse((string)paras[0]);
			int index = Convert.ToInt32(paras[1]);
			List<double> sizes = new List<double>();
			foreach (List<string> i in content.Series) {
				sizes.Add(double.Parse(i[2].Trim()));

			}
			//等之后可能性多了，改为map驱动
			C.DataLabel dataLabel = null;
			if (size < sizes.Max() / 4) {
				dataLabel = CreateDataLabel("", format["small"], 0);
			} else {
				dataLabel = CreateDataLabel("", format["big"], 0);
			}
			var xList = content.Series.ConvertAll<double>(x => double.Parse(x[0].Trim()));
			var yList = content.Series.ConvertAll<double>(x => double.Parse(x[1].Trim()));
			var rate = (yList.Max() - yList.Min()) / (xList.Max() - xList.Min());
			List<List<double>> intersectPoints = GetIntersectList(content.Series[index].ConvertAll<double>(x => Double.Parse(x.Trim())), 
				content.Series.ConvertAll<List<double>>(x => x.ConvertAll<double>(y => Double.Parse(y.Trim()))), rate, format);
			if(intersectPoints.Count > 1) {
				string nodeName = "position" + GetPosition(content.Series[index].ConvertAll<double>(x => Double.Parse(x.Trim())),
				intersectPoints, rate);
				dataLabel = SetPosition(dataLabel, format[nodeName]);
			}

			return dataLabel;
		}

		//读C.DataLabelPosition和layout进行改动
		private C.DataLabel SetPosition(C.DataLabel dataLabel, JToken format) {
			var manual = dataLabel.Elements<C.Layout>().First().Elements<C.ManualLayout>().First();
			manual.Elements<C.Left>().First().Val = Double.Parse((string)format["left"]);
			manual.Elements<C.Top>().First().Val = Double.Parse((string)format["top"]);
			dataLabel.Elements<C.DataLabelPosition>().First().Val = (C.DataLabelPositionValues)Enum.Parse(typeof(C.DataLabelPositionValues), (string)format["position"]);
			return dataLabel;
		}

		private List<List<double>> GetIntersectList(List<double> point, List<List<double>> series, double rate, JToken format) {

			List <List<double>> re = new List<List<double>>();
			foreach (List<double> j in series) {
				var lengthSquare = Math.Pow(point[0] * rate - j[0] * rate, 2) + Math.Pow(point[1] - j[1], 2);
				if (lengthSquare < double.Parse((string)format["diameterSquare"]) && lengthSquare > 0) {
					re.Add(j);
				}
			}
			return re;
		}

		private int GetPosition(List<double> point, List<List<double>> otherPoints, double rate) {
			List<double> angles = new List<double>();
			int[] positions = new int[4];
			foreach (List<double> otherPoint in otherPoints) {
				angles.Add(Math.Atan2(otherPoint[1] - point[1], otherPoint[0] * rate - point[0] * rate) * 180 / Math.PI);
			}
			foreach (double angle in angles) {
				var index = (int)(angle + 180) / 90;
				positions[index] += 1;
			}

			return positions.ToList().IndexOf(positions.Min());
		}

		private delegate C.DataLabel GetDataLabel(PhChartContent content, JToken format, params object[] paras);
	}
}
