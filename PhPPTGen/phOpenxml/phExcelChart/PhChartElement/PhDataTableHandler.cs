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
	class PhDataTableHandler : PhBaseElementHandler {
		protected override OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format) {
			C.DataTable dataTable = new C.DataTable();
			dataTable.Append(new C.ShowHorizontalBorder() { Val = true });
			dataTable.Append(new C.ShowVerticalBorder() { Val = true });
			dataTable.Append(new C.ShowOutlineBorder() { Val = true });
			dataTable.Append(new C.ShowKeys() { Val = true });
			return dataTable;
		}
	}
}
