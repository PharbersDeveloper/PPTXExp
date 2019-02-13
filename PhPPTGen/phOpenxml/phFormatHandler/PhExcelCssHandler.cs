using PhPPTGen.phCommand.phExcel.css;
using PhPPTGen.phModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phOpenxml.phFormatHandler {
	class PhExcelCssHandler {
		private static PhExcelCssHandler _instance = null;

		private PhExcelCssHandler() { }

		public static PhExcelCssHandler GetInstance() {
			if (_instance == null) {
				_instance = new PhExcelCssHandler();
			}
			return _instance;
		}

		public PhExcelCssForOpenxml Css2CellFormatName(string cssName) {
			Css.init();
			PhExcelCssForOpenxml css = new PhExcelCssForOpenxml() {
				fontSize = "",
				fontColor = "",
				fontName = "",
				bold = "",
				cellColor = "",
				height = "10.75",
				width = "10",
				verticalAlignType = "Center",
				horizontalAlignType = "Center",
				numbering = "",
			};
			foreach (string oneCss in cssName.Split('*')) {
				css = mergeCss(css, Css.getOpenxmlCss(oneCss));
			}
			return css;
		}

		protected PhExcelCssForOpenxml mergeCss(PhExcelCssForOpenxml oldCss, PhExcelCssForOpenxml newCss) {
			PropertyInfo[] propertys = oldCss.GetType().GetProperties();
			foreach (PropertyInfo property in propertys) {
				var newValue = newCss.GetType().GetProperty(property.Name).GetValue(newCss);
				if (newValue != null && !((string)newValue).Equals("")) {
					property.SetValue(oldCss, newValue);
				}
			}
			return oldCss;
		}
	}
}
