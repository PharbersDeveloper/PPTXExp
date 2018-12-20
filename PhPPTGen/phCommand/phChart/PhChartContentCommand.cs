using System;
using System.IO;
using Spire.Xls;
using Spire.Presentation;
using System.Drawing;
using Spire.Presentation.Drawing;
using System.Data;
using Spire.Presentation.Charts;
using PhPPTGen.phCommand.phExcel;

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
			string workbookKey = jobid + e2p.name;

			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2p.name + ".xls";
			if (!PhUpdateXlsCommand.workbookMap.ContainsKey(workbookKey)) {
				throw new Exception("Excel name is not exists");
			}

			Workbook book = new Workbook();
			PhUpdateXlsCommand.workbookMap.TryGetValue(workbookKey, out book);
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
            //chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 5;
            chart.ChartDataTable.Text.AutofitType = TextAutofitType.Normal;
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
            //chart.ChartDataTable.Text.
    
            ppt.SaveToFile(ppt_path, Spire.Presentation.FileFormat.Pptx2010);
            book.SaveToFile(ePath);
            ppt = new Presentation();
            ppt.LoadFromFile(ppt_path);
            foreach(Shape shape in ppt.Slides[e2p.slider].Shapes)
            {
                if(shape is IChart)
                {
                    chart = shape as IChart;
                    chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 6;
                }
            }
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
						chart.ChartData[i + 1, j].Value = Math.Round(number,2);
					} else {
						chart.ChartData[i +1, j].Value = s;
					}		
				}
				
			}
				
		}


        //static void Main(string[] args)
        //{
        //    phModel.PhRequest phRequest = new phModel.PhRequest();
        //    phModel.PhExcel2PPT phExcel2PPT = new phModel.PhExcel2PPT();
        //    phModel.PhExcelCss phExcelCss = new phModel.PhExcelCss()
        //    {
        //        cell = "A1",
        //        cellBordersColor = "#AEEEEE",
        //        cellBorders = new string[2] { "top#Thin", "bottom#Thin" },
        //        cellColor = "#000000"
        //    };
        //    phModel.PhExcelPush PhExcelPush = new phModel.PhExcelPush()
        //    {
        //        name = "testCss",
        //        cell = "A1",
        //        cate = "String",
        //        value = "test",
        //        css = phExcelCss
        //    };

        //    phExcel2PPT.name = "test";
        //    phExcel2PPT.slider = 1;
        //    phExcel2PPT.pos = new int[4] { 50, 60, 600, 400 };
        //    Workbook workbook = new Workbook();
        //    workbook.LoadFromFile(@"D:\pptresult\test\test.xls");
        //    PhUpdateXlsCommand.workbookMap.Add("testtest", workbook);
        //    phRequest.jobid = "test";
        //    phRequest.e2p = phExcel2PPT;
        //    phRequest.push = PhExcelPush;
        //    new PhCreatePPTCommand().Exec(phRequest);
        //    new PhChartContentCommand().Exec(phRequest);
        //    ////for(int i = 1; i < 20; i++) {
        //    ////	phExcelCss.cell = "A" + i;
        //    ////	PhExcelPush.cell = "A" + i;
        //    ////	new phExcel.PhUpdateXlsCommand().Exec(phRequest);
        //    ////}

        //    //Workbook workbook = new Workbook();
        //    //workbook.LoadFromFile(@"C:\Users\ycq\Documents\pptresult\test\testCss.xls");
        //    //Worksheet sheet = workbook.Worksheets[0];
        //    //for (int i = 1; i < 100; i++)
        //    //{
        //    //    phExcelCss.cell = "A" + i;
        //    //    new phExcel.PhSetXlsCssBaseCommand().Exec(phExcelCss, sheet);
        //    //}
        //    //workbook.SaveToFile(@"C:\Users\ycq\Documents\pptresult\test\testCss.xls");
        //}

    }
}
