using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhPieChart : Ph3DPieChart{
		protected override void DiyChart(IChart chart) {
			base.DiyChart(chart);
			chart.Type = ChartType.Pie;
			chart.Series[0].DataLabels.LeaderLinesVisible = true;
		}
	}
}
