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
	class PhCategoryAxisHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.CategoryAxis categoryAxis = new C.CategoryAxis();
			categoryAxis.Append(new C.AxisId() { Val = (UInt32Value)uint.Parse((string)format["axisId"]) });
			//方向，由format确定
			categoryAxis.Append(new C.Scaling(new C.Orientation() { Val = (C.OrientationValues)Enum.Parse(typeof(C.OrientationValues), (string)format["orientation"]) }));
			//是否显示轴
			categoryAxis.Append(new C.Delete() { Val = Boolean.Parse((string)format["delete"])});
			//轴位置
			categoryAxis.Append(new C.AxisPosition() { Val = (C.AxisPositionValues)Enum.Parse(typeof(C.AxisPositionValues), (string)format["axisPosition"]) });
			//编码方式，如保留小数等
			categoryAxis.Append(new C.NumberingFormat() { FormatCode = "General", SourceLinked = true });

			//刻度
			categoryAxis.Append(new C.MajorTickMark() { Val = (C.TickMarkValues)Enum.Parse(typeof(C.TickMarkValues), (string)format["majorTickMark"]) });
			categoryAxis.Append(new C.MinorTickMark() { Val = (C.TickMarkValues)Enum.Parse(typeof(C.TickMarkValues), (string)format["minorTickMark"]) });
			categoryAxis.Append(new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo });

			categoryAxis.Append(new C.CrossingAxis() { Val = (UInt32Value)uint.Parse((string)format["crossingAxis"]) });

			//轴交叉方式
			categoryAxis.Append(new C.Crosses() { Val = (C.CrossesValues)Enum.Parse(typeof(C.CrossesValues), (string)format["crosses"]) });

			categoryAxis.Append(new C.AutoLabeled() { Val = true });
			categoryAxis.Append(new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center });

			//轴上标签偏移
			categoryAxis.Append(new C.LabelOffset() { Val = (UInt16Value)uint.Parse((string)format["labelOffset"]) });

			categoryAxis.Append(new C.NoMultiLevelLabels() { Val = false });

			return categoryAxis;
		}
	}
}
