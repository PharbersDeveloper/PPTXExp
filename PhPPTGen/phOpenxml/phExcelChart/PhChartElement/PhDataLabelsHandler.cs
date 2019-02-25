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
	class PhDataLabelsHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.DataLabels dataLabels = new C.DataLabels();

			//标签属性，以后可以由format属性定义
			dataLabels.Append(new C.DataLabelPosition() { Val = C.DataLabelPositionValues.OutsideEnd });
			dataLabels.Append(new C.ShowLegendKey() { Val = false });
			dataLabels.Append(new C.ShowValue() { Val = true });
			dataLabels.Append(new C.ShowCategoryName() { Val = false });
			dataLabels.Append(new C.ShowSeriesName() { Val = false });
			dataLabels.Append(new C.ShowPercent() { Val = false });
			dataLabels.Append(new C.ShowBubbleSize() { Val = false });
			dataLabels.Append(new C.ShowLeaderLines() { Val = false });

			dataLabels.Append(CreateDLblsExtensionList());
			return dataLabels;
		}

		private C.DLblsExtensionList CreateDLblsExtensionList() {
			C.DLblsExtensionList dLblsExtensionList1 = new C.DLblsExtensionList();

			C.DLblsExtension dLblsExtension1 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
			dLblsExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
			C15.ShowLeaderLines showLeaderLines2 = new C15.ShowLeaderLines() { Val = true };

			C15.LeaderLines leaderLines1 = new C15.LeaderLines();

			C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

			A.Outline outline7 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

			A.SolidFill solidFill11 = new A.SolidFill();

			A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 35000 };
			A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 65000 };

			schemeColor20.Append(luminanceModulation12);
			schemeColor20.Append(luminanceOffset4);

			solidFill11.Append(schemeColor20);
			A.Round round1 = new A.Round();

			outline7.Append(solidFill11);
			outline7.Append(round1);
			A.EffectList effectList7 = new A.EffectList();

			chartShapeProperties4.Append(outline7);
			chartShapeProperties4.Append(effectList7);

			leaderLines1.Append(chartShapeProperties4);

			dLblsExtension1.Append(showLeaderLines2);
			dLblsExtension1.Append(leaderLines1);

			dLblsExtensionList1.Append(dLblsExtension1);
			return dLblsExtensionList1;
		}
	}
}
