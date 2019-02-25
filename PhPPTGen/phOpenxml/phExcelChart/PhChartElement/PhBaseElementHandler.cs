using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phExcelChart.DO;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	abstract class PhBaseElementHandler : IPhElementHandler {
		public OpenXmlCompositeElement CreateElement(PhChartContent content, JToken format) {
			var element = AppendDefaultElement(content, format);
			AppendChildElement(element, content, format);
			return element;
		}

		protected virtual void AppendChildElement(OpenXmlCompositeElement element, PhChartContent content, JToken format) {
			foreach(string childName in (JArray)format["child"]) {
				element.Append(GetHandler((string)format[childName]["factory"]).CreateElement(content, format[childName]));
			}
		}

		protected abstract OpenXmlCompositeElement AppendDefaultElement(PhChartContent content, JToken format);

		protected OpenXmlCompositeElement AppendOneElement(PhChartContent content, JToken format) {
			return GetHandler((string)format["factory"]).CreateElement(content, format);
		}

		protected IPhElementHandler GetHandler(string factory) {
			Assembly assembly = Assembly.GetExecutingAssembly();
			object obj = assembly.CreateInstance(factory);
			return (IPhElementHandler)obj;
		}
	}
}
