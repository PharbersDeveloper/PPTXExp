using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhColumnChart : PhChartBase {

		protected override void SetSeriesAndCategories(IChart chart, DataTable dt) {
			chart.Categories.CategoryLabels = chart.ChartData["A2", "A" + (dt.Rows.Count + 1)];
			chart.Series.SeriesLabel = chart.ChartData["B1", ((char)((int)'A' + (dt.Columns.Count - 1))).ToString() + "1"];
			for (int i = 0; i < dt.Columns.Count - 1; i++) {
				string start = ((char)((int)'B' + i)).ToString() + 2;
				string end = ((char)((int)'B' + i)).ToString() + (dt.Rows.Count + 1);
				chart.Series[i].Values = chart.ChartData[start, end];
			}
		}

		protected override void DiyChart(IChart chart) {
			chart.Type = ChartType.ColumnClustered;
			chart.HasDataTable = true;
			chart.HasTitle = false;
			chart.HasLegend = false;
			TextParagraph par = new TextParagraph();
			par.DefaultCharacterProperties.FontHeight = 10;
			//chart.ChartDataTable.Text.Paragraphs.Append(par);
		}
	}
}
