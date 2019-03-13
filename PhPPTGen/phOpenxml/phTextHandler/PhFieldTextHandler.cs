using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phOpenxml.phTextHandler {
	class PhFieldTextHandler: PhRunTextHandler {

		override public OpenXmlCompositeElement CreateRun(JToken fieldFormat) {
			A.Field field = new A.Field() { Id = (string)fieldFormat["id"], Type = (string)fieldFormat["fieldType"] };
			AppendContent(field, fieldFormat);
			//不知道smtClean的作用T.T
			field.Elements<A.RunProperties>().First().SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
			return field;
		}
	}
}
