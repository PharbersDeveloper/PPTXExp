using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phOpenxml.phTextHandler {
	class PhRunTextHandler {
		public virtual OpenXmlCompositeElement CreateRun(JToken runFormat) {
			A.Run run = new A.Run();
			AppendContent(run, runFormat);
			return run;
		}

		protected void AppendContent(OpenXmlCompositeElement element, JToken format) {
			A.Text text = new A.Text() { Text = (string)format["text"] };
			A.RunProperties properties = new A.RunProperties() {
				Language = "en-US",
				AlternativeLanguage = "zh-CN",
				FontSize = int.Parse((string)format["fontSize"]) * 100,
				Bold = Boolean.Parse((string)format["bold"]),
				Dirty = false
			};
			A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex() { Val = (string)format["color"] });
			properties.Append(solidFill);
			element.Append(properties, text);
		}
	}
}
