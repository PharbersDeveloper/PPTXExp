using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phChart {
	//以后可以通过配置文件实现
	class PhChartType {
		public static string line = "Line";
		public static string combo = "Combo";
		public static string lineNoTable = "LineNoTable";


		public static string PhChaerType2Cls(string type) {
			if (type == line) {
				return "PhPPTGen.phCommand.phChart.PhLineChart";
			} else if (type == combo) {
				return "PhPPTGen.phCommand.phChart.PhComboChart";
			} else if (type == lineNoTable) {
				return "PhPPTGen.phCommand.phChart.PhLineChartNoTable";
			} else {
				throw new System.Exception("Can not handler message");
			}
		}
	}
}
