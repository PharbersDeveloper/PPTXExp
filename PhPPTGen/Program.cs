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
		//static void Main(string[] args) {
		//	phSocket.PhThreadSocketServ s = new phSocket.PhThreadSocketServ();
		//	s.startListen();
		//	phCommon.PhMsgLst lst = phCommon.PhMsgLst.GetInstance();
		//	lst.StartChecking();
		//}

		static void Main(string[] args) {
			PhChartPartsHandler test = new PhChartPartsHandler();
			
			using (StreamReader reader = File.OpenText(@"..\..\resources\PhBarChart.json")) {
				test.Format = JToken.ReadFrom(new JsonTextReader(reader))["chart"];

			}
			using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(@"D:\alfredyang\chartTest.xlsx", true)) {
				WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;
				test.Content.SetValueFromExcel(workbookPart);
				WorksheetPart worksheetPart = mySpreadsheet.WorkbookPart.WorksheetParts.ElementAt(0);
				DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
				worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
				worksheetPart.Worksheet.Save();
				ChartPart chartPart1 = drawingsPart.AddNewPart<ChartPart>("ch1");
				test.CreateChartPart(chartPart1);
				into(drawingsPart);
				chartPart1.ChartSpace.Save();
				drawingsPart.WorksheetDrawing.Save();
			}
		}

		//测试用代码
		private static void into(DrawingsPart drawingsPart) {
			drawingsPart.WorksheetDrawing = new A.Spreadsheet.WorksheetDrawing();
			A.Spreadsheet.TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<A.Spreadsheet.TwoCellAnchor>(new A.Spreadsheet.TwoCellAnchor());
			twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new A.Spreadsheet.ColumnId("9"),
				new A.Spreadsheet.ColumnOffset("581025"),
				new A.Spreadsheet.RowId("17"),
				new A.Spreadsheet.RowOffset("114300")));
			twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new A.Spreadsheet.ColumnId("17"),
				new A.Spreadsheet.ColumnOffset("276225"),
				new A.Spreadsheet.RowId("32"),
				new A.Spreadsheet.RowOffset("0")));

			// Append a GraphicFrame to the TwoCellAnchor object.
			DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
				twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
				Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
				Spreadsheet.GraphicFrame());
			graphicFrame.Macro = "";

			graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
				new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
				new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

			graphicFrame.Append(new A.Spreadsheet.Transform(new A.Offset() { X = 0L, Y = 0L },
																	new A.Extents() { Cx = 0L, Cy = 0L }));

			graphicFrame.Append(new A.Graphic(new A.GraphicData(new C.ChartReference() { Id = "ch1" }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

			twoCellAnchor.Append(new A.Spreadsheet.ClientData());

			// Save the WorksheetDrawing object.
			drawingsPart.WorksheetDrawing.Save();
		}
	}
}