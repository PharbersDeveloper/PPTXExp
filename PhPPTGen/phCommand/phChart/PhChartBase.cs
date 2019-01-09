using System;
using System.IO;
using Spire.Xls;
using Spire.Presentation;
using System.Drawing;
using Spire.Presentation.Drawing;
using System.Data;
using Spire.Presentation.Charts;
using PhPPTGen.phCommand.phExcel;
using System.Collections.Generic;

namespace PhPPTGen.phCommand.phChart {
	class PhChartBase : PhCommand {
		public override object Exec(params object[] parameters) {
			PutChart(parameters);
			return null;
		}

		protected virtual void PutChart(params object[] parameters) {
			Console.WriteLine("PPTGen phCommand: generate chart shape for ppt");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var e2c = req.e2c;
			/**
             * 1. go to the work place
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;
			string workbookKey = jobid + e2c.name;

			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2c.name + ".xls";
			if (!PhUpdateXlsCommand.workbookMap.ContainsKey(workbookKey)) {
				throw new Exception("Excel name is not exists");
			}

			Workbook book = new Workbook();
			PhUpdateXlsCommand.workbookMap.TryGetValue(workbookKey, out book);
            PhUpdateXlsCommand.workbookMap.Remove(workbookKey);
            Worksheet sheet = book.Worksheets[0];
			var col = sheet.Columns.Length;
			var row = sheet.Rows.Length;

			/**
             * 3.读取workbook到datatable
             */
			Spire.Xls.CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
			DataTable dt = sheet.ExportDataTable(range, true, true);

			/**
             * 4.插入图表
             */
			String ppt_path = workingPath + "\\" + "result.pptx";

			Presentation ppt = new Presentation();
			ppt.LoadFromFile(ppt_path);

			while (ppt.Slides.Count <= e2c.slider) {
				ppt.Slides.Append();
			}
			Rectangle rec = new Rectangle(e2c.pos[0], e2c.pos[1], e2c.pos[2], e2c.pos[3]);
			IChart chart = ppt.Slides[e2c.slider].Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Line, rec);
			InitChartData(chart, dt);
			SetSeriesAndCategories(chart, dt);

			chart.HasDataTable = true;
			DiyChart(chart);
			ppt.SaveToFile(ppt_path, Spire.Presentation.FileFormat.Pptx2010);
			book.SaveToFile(ePath);
			//Presentation pptx = new Presentation();
			//pptx.LoadFromFile(ppt_path);
			//foreach (Shape shape in pptx.Slides[e2c.slider].Shapes) {
			//	if (shape is IChart) {
			//		chart = shape as IChart;
			//		DiyChart(chart);
			//	}
			//}
			//pptx.SaveToFile(ppt_path, Spire.Presentation.FileFormat.Pptx2010);
		}

		protected virtual void SetSeriesAndCategories(IChart chart, DataTable dt) {
			chart.Series.SeriesLabel = chart.ChartData["A2", "A" + (dt.Rows.Count + 1)];
			chart.Categories.CategoryLabels = chart.ChartData["B1", ((char)((int)'A' + (dt.Columns.Count - 1))).ToString() + "1"];
			for (int i = 0; i < dt.Rows.Count; i++) {
				string start = "B" + (i + 2);
				string end = ((char)((int)'A' + (dt.Columns.Count - 1))).ToString() + (i + 2);
				chart.Series[i].Values = chart.ChartData[start, end];
			}
		}

		protected virtual void DiyChart(IChart chart) {
			chart.HasTitle = false;
			chart.HasLegend = false;
			chart.ChartDataTable.ShowLegendKey = true;
			chart.ChartDataTable.Text.AutofitType = TextAutofitType.Normal;
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			chart.PrimaryValueAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
            TextParagraph par = new TextParagraph();
            par.DefaultCharacterProperties.FontHeight = 8;
			chart.ChartDataTable.Text.Paragraphs.Append(par);
            //chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
        }

        protected virtual void InitChartData(IChart chart, DataTable dataTable) {
            for (int c = 0; c < dataTable.Columns.Count; c++) {
                chart.ChartData[0, c].Text = dataTable.Columns[c].Caption;
            }

            //for (int r = 0; r < dataTable.Rows.Count; r++)

            //{
            //    object[] data = dataTable.Rows[r].ItemArray;
            //    for(int c = 0; c < data.Length; c++)
            //    {
            //        chart.ChartData[r + 1,c].Value = (int)data[c];
            //    }
            //}



            for (int i = 0; i < dataTable.Rows.Count; i++) {
                for (int j = 0; j < dataTable.Rows[0].ItemArray.Length; j++) {
                    Double number = 0;
                    string s = dataTable.Rows[i].ItemArray[j].ToString();
                    bool result = Double.TryParse(s, out number);
                    if (result) {
                        chart.ChartData[i + 1, j].Value = Math.Round(number, 2);
                    } else {
                        if (s.Equals("N/A")) {
                            chart.ChartData[i + 1, j].Value = 100;
                        } else {
                            chart.ChartData[i + 1, j].Value = s;
                        }

                    }
                }

            }
        }
	}
}
