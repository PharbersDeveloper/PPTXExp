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
            foreach (var i in chart.ChartLegend.EntryTextProperties) {
                i.FontHeight = 8;
            }
            chart.ChartLegend.Position = ChartLegendPositionType.Top;
            chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
            chart.PrimaryValueAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
            chart.HasDataTable = false;
		}
	}
}
