﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhComboChart : PhChartContentCommand {
		public override object Exec(params object[] parameters) {
			PutChart(parameters);
			return null;
		}

		protected override void DiyChart(IChart chart) {
			chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			chart.ChartLegend.Position = ChartLegendPositionType.Top;
			chart.Series[0].Type = ChartType.ColumnClustered;
			chart.Series[0].Fill.SolidColor.KnownColor = Spire.Presentation.KnownColors.Red;
			chart.Series[0].InvertIfNegative = false;
			for (int i = 0; i < chart.Series[0].Values.Count; i++) {
				ChartDataLabel lebel = chart.Series[0].DataLabels.Add();
				lebel.LabelValueVisible = true;
				lebel.PercentageVisible = true;
				//lebel.NumberFormat = "#,##0.0%";
				lebel.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10.5f;
				//lebel.TextProperties.Paragraphs[0].Text = string.Format("{0:P}", lebel.TextProperties.Paragraphs[0].Text);
				//lebel.TextFrame.Text;
				//lebel.Position = ChartDataLabelPosition.InsideEnd;

			}
			for(int i = 0; i < chart.Series[1].Values.Count; i++) {
				ChartDataLabel lebel = chart.Series[1].DataLabels.Add();
				lebel.LabelValueVisible = i == (chart.Series[1].Values.Count - 1);
				lebel.PercentageVisible = true;
				lebel.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 12;
				lebel.X = 10f;
				lebel.Y = 9f;
			}
			chart.Series[1].Type = ChartType.Line;
			chart.Series[1].Line.DashStyle = Spire.Presentation.LineDashStyleType.Dash;
			ChartDataPoint cdp = new ChartDataPoint(chart.Series[1]);

			chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow;
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			chart.PrimaryValueAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			//chart.PrimaryValueAxis.NumberFormat = "#,##0.0%";
		}
	}
}
