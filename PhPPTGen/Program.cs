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

namespace PhPPTGen {
	class Program {
		static void Main(string[] args) {
			phSocket.PhThreadSocketServ s = new phSocket.PhThreadSocketServ();
			s.startListen();
			phCommon.PhMsgLst lst = phCommon.PhMsgLst.GetInstance();
			lst.StartChecking();
		}

		//static void Main(string[] args) {
		//	PhChartPartsHandler test = new PhChartPartsHandler();

		//	using (StreamReader reader = File.OpenText(@"..\..\resources\PhBarChart.json")) {
		//		test.Format = JToken.ReadFrom(new JsonTextReader(reader))["chart"];

		//	}
		//	using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(@"D:\alfredyang\chartTest.xlsx", true)) {
		//		WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;
		//		test.Content.SetValueFromExcel(workbookPart);
		//		WorksheetPart worksheetPart = mySpreadsheet.WorkbookPart.WorksheetParts.ElementAt(0);
		//		DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
		//		worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
		//		worksheetPart.Worksheet.Save();
		//		ChartPart chartPart1 = drawingsPart.AddNewPart<ChartPart>("ch1");
		//		test.CreateChartPart(chartPart1);
		//		into(drawingsPart);
		//		chartPart1.ChartSpace.Save();
		//		drawingsPart.WorksheetDrawing.Save();
		//	}
		//}


	}
}