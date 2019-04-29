using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PhPPTGen.phModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phExcel.css {
    class Css {
        private static Dictionary<string, PhExcelCss> cssMap = new Dictionary<string, PhExcelCss>();
		private static Dictionary<string, PhExcelCssForOpenxml> openxmlCssMap = new Dictionary<string, PhExcelCssForOpenxml>();
		public static void init() {
            if (cssMap.Count == 0) {
                string json = @"{}";
				//openxmlCssMap = JsonConvert.DeserializeObject<Dictionary<string, PhExcelCssForOpenxml>>(xmlJson);
				cssMap = JsonConvert.DeserializeObject <Dictionary<string, PhExcelCss>>(json);
				using (StreamReader reader = File.OpenText(PhConfigHandler.GetInstance().path + PhConfigHandler.GetInstance().GetConfigMap()["excelCss"].Value<string>())) {
					var cssFormat = JToken.ReadFrom(new JsonTextReader(reader));
					openxmlCssMap = JsonConvert.DeserializeObject<Dictionary<string, PhExcelCssForOpenxml>>(cssFormat.ToString());
				}
			}
        }
        public static PhExcelCss getCss(string cssName) {
            PhExcelCss css = new PhExcelCss();

            if (cssMap.TryGetValue(cssName, out css)) {
                return css;
            }
            return new PhExcelCss();
        }

		public static PhExcelCssForOpenxml getOpenxmlCss(string cssName) {
			PhExcelCssForOpenxml css = new PhExcelCssForOpenxml();

			if (openxmlCssMap.TryGetValue(cssName, out css)) {
				return css;
			}
			return new PhExcelCssForOpenxml();
		}
	}
}
