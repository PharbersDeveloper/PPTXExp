using System;
using System.IO;
using Spire.Xls;
using Spire.Presentation;
using System.Drawing;
using Spire.Presentation.Drawing;
using System.Data;
using Spire.Presentation.Charts;

namespace PhPPTGen.phCommand.phChart {
	class PhChartContentCommand : PhCommand {
		public override object Exec(params object[] parameters) {
			Console.WriteLine("PPTGen phCommand: generate chart shape for ppt");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var e2p = req.e2p;
			/**
             * 1. go to the work place
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;

			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2p.name + ".xls";
			if (!File.Exists(ePath)) {
				throw new Exception("Excel name is not exists");
			}

			Workbook book = new Workbook();
			book.LoadFromFile(ePath);
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

			while (ppt.Slides.Count <= e2p.slider) {
				ppt.Slides.Append();
			}
			Rectangle rec = new Rectangle(e2p.pos[0], e2p.pos[1], e2p.pos[2], e2p.pos[3]);
			IChart chart = ppt.Slides[e2p.slider].Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Line, rec);
			//chart.ChartTitle.TextProperties.Text = "部门信息";
			//chart.ChartTitle.TextProperties.IsCentered = true;
			//chart.ChartTitle.Height = 30;
			chart.HasTitle = false;
			chart.HasLegend = false;
			chart.ChartDataTable.ShowLegendKey = true;
			chart.HasDataTable = true;
			chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			chart.PrimaryValueAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 8;
			//chart.PlotArea.Top = 1;
			//chart.PlotArea.Left = 100;

			//for (int i = 0; i < data.GetLength(0); i++)
			//{
			//    for (int j = 0; j < data.GetLength(1); j++)
			//    {
			//        //将数字类型的字符串转换为整数
			//        int number;
			//        bool result = Int32.TryParse(data[i, j], out number);
			//        if (result)
			//        {
			//            chart.ChartData[i, j].Value = number;
			//        }
			//        else
			//        {
			//            chart.ChartData[i, j].Value = data[i, j];
			//        }
			//    }
			//}
			InitChartData(chart, dt);
			chart.Series.SeriesLabel = chart.ChartData["A2", "A" + row];
			chart.Categories.CategoryLabels = chart.ChartData["B1", ((char)((int)'A' + (dt.Columns.Count - 1))).ToString() + "1"];
			for (int i = 0; i < dt.Rows.Count; i++) {
				string start = "B" + (i + 2);
				string end = ((char)((int)'A' + (dt.Columns.Count - 1))).ToString() + (i + 2);
				chart.Series[i].Values = chart.ChartData[start, end];
			}
			//chart.ChartStyle = ChartStyle.Style11;
			//chart.GapWidth = 200;

			ppt.SaveToFile(ppt_path, Spire.Presentation.FileFormat.Pptx2010);

			return null;
		}

		private void InitChartData(IChart chart, DataTable dataTable) {
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
				for (int j = 0; j < dataTable.Rows[0].ItemArray.Length ; j++) {
					Double number;
					string s = dataTable.Rows[i].ItemArray[j].ToString();
					bool result = Double.TryParse(s, out number);
					if (result) {
						chart.ChartData[i + 1, j].Value = number;
					} else {
						chart.ChartData[i +1, j].Value = s;
					}		
				}
				
			}
				
		}


		static void Main(string[] args) {
			phModel.PhRequest phRequest = new phModel.PhRequest();
			phModel.PhExcel2PPT phExcel2PPT = new phModel.PhExcel2PPT();
			phModel.PhExcelCss phExcelCss = new phModel.PhExcelCss() {
				cell = "A1", cellBordersColor = "#F5F5F5",
				cellBorders =new string[2] {"top#Thin", "bottom#Thin"}
			};
			phModel.PhExcelPush PhExcelPush = new phModel.PhExcelPush() {
				name = "testCss", cell = "A1", cate = "String", value = "test",
				css = phExcelCss
			};
			
			phExcel2PPT.name = "test";
			phExcel2PPT.slider = 1;
			phExcel2PPT.pos = new int[4] { 50, 60, 600, 200 };
	
		

			phRequest.jobid = "test";
			phRequest.e2p = phExcel2PPT;
			phRequest.push = PhExcelPush;
			new PhCreatePPTCommand().Exec(phRequest);
			//new PhChartContentCommand().Exec(phRequest);
			new phExcel.PhUpdateXlsCommand().Exec(phRequest);
		}

	}
}
