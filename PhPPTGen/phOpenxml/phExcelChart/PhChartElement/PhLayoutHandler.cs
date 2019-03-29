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
	class PhLayoutHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.Layout layout = new C.Layout();

			C.ManualLayout manualLayout = new C.ManualLayout();
			C.LeftMode leftMode = new C.LeftMode() { Val = (C.LayoutModeValues)Enum.Parse(typeof(C.LayoutModeValues), (string)format["leftMode"]) };
			C.TopMode topMode = new C.TopMode() { Val = (C.LayoutModeValues)Enum.Parse(typeof(C.LayoutModeValues), (string)format["topMode"]) };
			C.Left left = new C.Left() { Val = Double.Parse((string)format["left"]) };
			C.Top top = new C.Top() { Val = Double.Parse((string)format["top"]) };

			manualLayout.Append(leftMode);
			manualLayout.Append(topMode);
			manualLayout.Append(left);
			manualLayout.Append(top);

			layout.Append(manualLayout);
			return layout;
		}
	}
}
