using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phOpenxml.PhShapeTreeChilren {
	interface IPhShapeTreeChilrenHander {
		OpenXmlElement CreateElement(JToken format, params object[] paras);
	}
}
