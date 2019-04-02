using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;

namespace PhPPTGen.phOpenxml.PhShapeTreeChilren {
	abstract class PhBaseHandler: IPhShapeTreeChilrenHander {
		public OpenXmlElement CreateElement(JToken format, params object[] paras) {
			var element = AppendDefaultElement(format, paras);
			AppendChildElement(element, format, paras);
			return element;
		}

		abstract protected OpenXmlElement AppendDefaultElement(JToken format, params object[] paras);

		protected virtual void AppendChildElement(OpenXmlElement element, JToken format, params object[] paras) {
			foreach (string childName in (JArray)format["chilren"]) {
				element.Append(GetHandler((string)format[childName]["factory"]).CreateElement(format[childName], paras));
			}
		}

		protected IPhShapeTreeChilrenHander GetHandler(string factory) {
			Assembly assembly = Assembly.GetExecutingAssembly();
			object obj = assembly.CreateInstance(factory);
			return (IPhShapeTreeChilrenHander)obj;
		}

	}
}
