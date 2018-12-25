using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class Ph3DPieChart : PhChartBase {

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
			chart.Type = ChartType.Pie3D;
			chart.HasDataTable = false;
			chart.HasTitle = false;
			chart.HasLegend = false;
			for (int i = 0; i < chart.Series[0].Values.Count; i++) {
				ChartDataLabel lebel = chart.Series[0].DataLabels.Add();
				lebel.CategoryNameVisible = true;
				lebel.LabelValueVisible = false;
				lebel.PercentageVisible = true;
				lebel.NumberFormat = "0.00";
				lebel.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10.5f;
				//lebel.Position = ChartDataLabelPosition.OutsideEnd;

			}
		}
	}
}
