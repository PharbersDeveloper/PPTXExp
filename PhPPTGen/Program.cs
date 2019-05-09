using PhPPTGen.phOpenxml.phExcelChart.PhChartElement;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Spire.Xls;
using PhPPTGen.phkafka;

namespace PhPPTGen {
	class Program {
		static void Main(string[] args) {
			var msg = File.OpenText(@"D:\alfredyang\msg.txt").ReadToEnd();
			PhConsumer.PushMsg(msg, "test");
			PhConfigHandler.GetInstance().path = @"..\..\resources\";
			phCommandFactory.PhCommandFactory.GetInstance().SeTmpDictionary(@"D:\pptresult\css\");
			PhConfigHandler.GetInstance().init();
			phCommon.PhRequestLst lst = phCommon.PhRequestLst.GetInstance();
			lst.StartChecking();
			PhConsumer.PullMsg("test");
		}

		//static void Main(string[] args) {
		//	PhConfigHandler.GetInstance().path = @"..\..\resources\";
		//	phCommandFactory.PhCommandFactory.GetInstance().SeTmpDictionary(@"D:\pptresult\css\");
		//	PhConfigHandler.GetInstance().init();

		//	//using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(@"D:\alfredyang\aa6452e1-bf63-47ab-bdf2-b5e19a5200c7.xlsx", true)) {
		//	//	//Get all the appropriate parts
		//	//	WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;

		//	//	//生成chart在excel中
		//	//	phOpenxml.PhExcelHandler.GetInstance().InsertChartIntoExcel(workbookPart, "Line");
		//	//}
		//	var input = new phModel.PhRequest() {
		//		jobid = "dcsTest",
		//		e2c = new phModel.PhExcel2Chart() {
		//			name = "test4 - 副本 - 副本",
		//			pos = new int[4] { (int)(254 / 0.000278), (int)(432 / 0.000278), (int)(1947 / 0.000278), (int)(1284 / 0.000278) },
		//			chartType = "Column",
		//			slider = 1
		//		}

		//		//e2c = new phModel.PhExcel2Chart() {
		//		//	name = "test",
		//		//	pos = new int[4] { (int)(169 / 0.000278), (int)(624 / 0.000278), (int)(927 / 0.000278), (int)(1105 / 0.000278) },
		//		//	chartType = "Stacked",
		//		//	slider = 1
		//		//}
		//	};
		//	new phCommand.phChart.PhPPTImportChartCommand().Exec(input);

		//}
	}
}