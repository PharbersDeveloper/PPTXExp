using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phChart {
	class PhLineChart : PhChartContentCommand {
		public override object Exec(params object[] parameters) {
			PutChart(parameters);
			return null;
		}
	}
}
