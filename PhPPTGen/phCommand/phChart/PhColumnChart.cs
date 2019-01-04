using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhColumnChart : PhChartBase {
		protected override void DiyChart(IChart chart) {
			chart.Type = ChartType.ColumnClustered;
			chart.HasDataTable = true;
			chart.HasTitle = false;
			chart.HasLegend = false;
			TextParagraph par = new TextParagraph();
			par.DefaultCharacterProperties.FontHeight = 10;
			chart.ChartDataTable.Text.Paragraphs.Append(par);
		}
	}
}
