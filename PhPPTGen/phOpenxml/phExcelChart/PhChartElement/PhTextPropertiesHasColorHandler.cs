using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	class PhTextPropertiesHasColorHandler: PhTextPropertiesHandler {
		protected override DefaultRunProperties GetDefaultRunProperties(JToken format) {
			A.DefaultRunProperties defaultRunProperties = base.GetDefaultRunProperties(format);
			A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex() { Val = (string)format["color"] });
			defaultRunProperties.Append(solidFill);
			return defaultRunProperties;
		}
	}
}
