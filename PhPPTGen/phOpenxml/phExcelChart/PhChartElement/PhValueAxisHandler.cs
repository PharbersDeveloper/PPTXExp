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
	class PhValueAxisHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.ValueAxis valueAxis = new C.ValueAxis();
			valueAxis.Append(new C.AxisId() { Val = (UInt32Value)uint.Parse((string)format["axisId"]) });
			valueAxis.Append(new C.Scaling(new C.Orientation() { Val = (C.OrientationValues)Enum.Parse(typeof(C.OrientationValues), (string)format["orientation"]) }));
			//方向，由format确定
			valueAxis.Append(new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }));
			//是否显示轴
			valueAxis.Append(new C.Delete() { Val = Boolean.Parse((string)format["delete"]) });
			//轴位置
			valueAxis.Append(new C.AxisPosition() { Val = (C.AxisPositionValues)Enum.Parse(typeof(C.AxisPositionValues), (string)format["axisPosition"]) });
			valueAxis.Append(new C.MajorGridlines(AppendOneElement(content, format["majorGridlinesShapeProperties"])));
			//编码方式，如保留小数等
			valueAxis.Append(new C.NumberingFormat() { FormatCode = (string)format["code"], SourceLinked = Boolean.Parse((string)format["sourceLinked"]) });

			//刻度
			valueAxis.Append(new C.MajorTickMark() { Val = (C.TickMarkValues)Enum.Parse(typeof(C.TickMarkValues), (string)format["majorTickMark"]) });
			valueAxis.Append(new C.MinorTickMark() { Val = (C.TickMarkValues)Enum.Parse(typeof(C.TickMarkValues), (string)format["minorTickMark"]) });
			valueAxis.Append(new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo });

			valueAxis.Append(new C.CrossingAxis() { Val = (UInt32Value)uint.Parse((string)format["crossingAxis"]) });

			//轴交叉方式
			valueAxis.Append(new C.Crosses() { Val = (C.CrossesValues)Enum.Parse(typeof(C.CrossesValues), (string)format["crosses"]) });
			valueAxis.Append(new C.CrossBetween() { Val = (C.CrossBetweenValues)Enum.Parse(typeof(C.CrossBetweenValues), (string)format["crossBetween"]) });

			return valueAxis;
		}
	}
}
