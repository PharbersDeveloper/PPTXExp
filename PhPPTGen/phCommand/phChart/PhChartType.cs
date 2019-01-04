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
		public static string Pie3D = "Pie3D";
		public static string Column = "Column";
		public static string Pie = "Pie";


		public static string PhChaerType2Cls(string type) {
			if (type == line) {
				return "PhPPTGen.phCommand.phChart.PhLineChart";
			} else if (type == combo) {
				return "PhPPTGen.phCommand.phChart.PhComboChart";
			} else if (type == lineNoTable) {
				return "PhPPTGen.phCommand.phChart.PhLineChartNoTable";
			} else if (type == Pie3D) {
				return "PhPPTGen.phCommand.phChart.Ph3DPieChart";
			} else if (type == Column) {
				return "PhPPTGen.phCommand.phChart.PhColumnChart";
			} else if (type == Pie) {
				return "PhPPTGen.phCommand.phChart.PhPieChart";
			} else {
				throw new System.Exception("Can not handler message");
			}
		}
	}
}
