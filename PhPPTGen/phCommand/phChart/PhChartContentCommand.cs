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
	class PhChartContentCommand : PhCommand {
		static Dictionary<string, string> workbookMap = new Dictionary<string, string>();
		public override object Exec(params object[] parameters) {
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var e2c = req.e2c;
			string ins = PhChartType.PhChaerType2Cls(e2c.chartType);
			phCommandFactory.PhCommandFactory fct = phCommandFactory.PhCommandFactory.GetInstance();
			fct.CreateCommandInstance(ins, parameters);
			return null;
		}


		//static void Main(string[] args) {
		//	phModel.PhRequest phRequest = new phModel.PhRequest();
		//	phModel.PhExcel2Chart phExcel2Chart = new phModel.PhExcel2Chart();
		//	phModel.PhExcelCss phExcelCss = new phModel.PhExcelCss() {
		//		cell = "A1",
		//		cellBordersColor = "#AEEEEE",
		//		cellBorders = new string[2] { "top#Thin", "bottom#Thin" },
		//		cellColor = "#000000"
		//	};
		//	phModel.PhExcelPush PhExcelPush = new phModel.PhExcelPush() {
		//		name = "testCss",
		//		cell = "A1",
		//		cate = "String",
		//		value = "test",
		//		css = phExcelCss
		//	};

		//	phExcel2Chart.name = "test";
		//	phExcel2Chart.slider = 1;
		//	phExcel2Chart.pos = new int[4] { 50, 60, 600, 400 };
		//	phExcel2Chart.chartType = "Pie3D";
		//	Workbook workbook = new Workbook();
		//	workbook.LoadFromFile(@"C:\Users\ycq\Documents\pptresult\test\test1.xls");
		//	PhUpdateXlsCommand.workbookMap.Add("testtest", workbook);
		//	phRequest.jobid = "test";
		//	phRequest.e2c = phExcel2Chart;
		//	phRequest.push = PhExcelPush;
		//	new PhCreatePPTCommand().Exec(phRequest);
		//	new PhChartContentCommand().Exec(phRequest);
		//	////for(int i = 1; i < 20; i++) {
		//	////	phExcelCss.cell = "A" + i;
		//	////	PhExcelPush.cell = "A" + i;
		//	////	new phExcel.PhUpdateXlsCommand().Exec(phRequest);
		//	////}

		//	//Workbook workbook = new Workbook();
		//	//workbook.LoadFromFile(@"C:\Users\ycq\Documents\pptresult\test\testCss.xls");
		//	//Worksheet sheet = workbook.Worksheets[0];
		//	//for (int i = 1; i < 100; i++)
		//	//{
		//	//    phExcelCss.cell = "A" + i;
		//	//    new phExcel.PhSetXlsCssBaseCommand().Exec(phExcelCss, sheet);
		//	//}
		//	//workbook.SaveToFile(@"C:\Users\ycq\Documents\pptresult\test\testCss.xls");
		//	//}

		//}
	}
}
