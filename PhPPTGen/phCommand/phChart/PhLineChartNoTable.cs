using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhLineChartNoTable : PhChartBase {
		protected override void DiyChart(IChart chart) {
			chart.HasTitle = false;
			chart.HasLegend = true;
			chart.ChartDataTable.ShowLegendKey = true;
			//chart.ChartDataTable.Text.AutofitType = TextAutofitType.Normal;
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			chart.HasDataTable = false;
		}
	}
}
