using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;
using DocumentFormat.OpenXml;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhView3DHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.View3D view3D = new C.View3D();
			C.RotateX rotateX1 = new C.RotateX() { Val = SByte.Parse((string)format["RotateX"]) };
			C.RotateY rotateY1 = new C.RotateY() { Val = UInt16.Parse((string)format["RotateY"]) };
			C.DepthPercent depthPercent1 = new C.DepthPercent() { Val = (UInt16Value)100U };
			C.RightAngleAxes rightAngleAxes1 = new C.RightAngleAxes() { Val = false };

			view3D.Append(rotateX1);
			view3D.Append(rotateY1);
			view3D.Append(depthPercent1);
			view3D.Append(rightAngleAxes1);
			return view3D;
		}
	}
}
