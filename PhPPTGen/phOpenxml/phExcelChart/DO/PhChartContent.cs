using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phOpenxml.phExcelChart.DO {
	class PhChartContent {
		public Dictionary<string, string> titleMap = new Dictionary<string, string>() { { "chartTitle", "chartTitle" } };
		public List<List<string>> Series = new List<List<string>>();
		public List<string> CategoryLabels = new List<string>();
		public List<string> seriesLabels = new List<string>();

		//public void SetValue() {
		//	Series = new List<string[]> { new string[3] { "1", "2", "3" }, new string[3] { "4", "5", "6" } };
		//	CategoryLabels = new string[2] { "1号", "2号" };
		//	seriesLabels = new string[3] { "a", "b", "c" };
		//}

		public void SetValueFromExcel(WorkbookPart workbookPart) {
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);
			SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
			var rows = sheetData.Elements<Row>().ToList();
			foreach (Cell c in rows.First()) {
				seriesLabels.Add(GetValue(c, workbookPart));
			}

			rows.Remove(rows.First());

			foreach (Row r in rows) {
				var cells = r.Elements<Cell>().ToList();
				CategoryLabels.Add(GetValue(cells.First(), workbookPart));
				cells.Remove(cells.First());
				List<string> serise = new List<string>();
				foreach (Cell c in cells) {
					serise.Add(GetValue(c, workbookPart));
				}
				Series.Add(serise);
			}
		}
		

		private string GetValue(Cell cell, WorkbookPart workbookPart) {
			GetCellValue nullTypeValue = (c, w) => c.InnerText;
			GetCellValue stringTypeValue = (c, w) => w.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
				.ElementAt(int.Parse(c.InnerText)).InnerText;

			Dictionary<CellValues, GetCellValue> factionMap = new Dictionary<CellValues, GetCellValue>() {
				{CellValues.SharedString, stringTypeValue }
			};
			if(cell.DataType == null) {
				return nullTypeValue(cell, workbookPart);
			}
			return factionMap[cell.DataType.Value](cell, workbookPart);
		}


		//private string GetNullTypeValue(Cell cell, WorkbookPart workbookPart) {
		//	return cell.InnerText;
		//}

		//private string GetSharedStringValue(Cell cell, WorkbookPart workbookPart) {
		//	return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
		//		.ElementAt(int.Parse(cell.InnerText)).InnerText;
		//}

		private delegate string GetCellValue(Cell c, WorkbookPart w);
	}

	
}
