using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
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

		public void SetValueFromExcel(WorkbookPart workbookPart, JToken format) {
			Dictionary<string, SetValue> funcMap = new Dictionary<string, SetValue>() {
				{"row", SetValueForRowType },{"column", SetValueForColumnType}
			};

			funcMap[(string)format["contentType"]](workbookPart);
		}
		

		private void SetValueForRowType(WorkbookPart workbookPart) {
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

		private void SetValueForColumnType(WorkbookPart workbookPart) {
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);
			SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
			var rows = sheetData.Elements<Row>().ToList();
			foreach (Cell c in rows.First()) {
				CategoryLabels.Add(GetValue(c, workbookPart));
			}

			rows.Remove(rows.First());

			for(int i = 0; i < rows.First().Elements<Cell>().Count() - 1; i++) {
				Series.Add(new List<string>());
			}
				foreach (Row r in rows) {
				var cells = r.Elements<Cell>().ToList();
				seriesLabels.Add(GetValue(cells.First(), workbookPart));
				cells.Remove(cells.First());
				List<string> serise = new List<string>();
				foreach (Cell c in cells) {
					Series[cells.IndexOf(c)].Add(GetValue(c, workbookPart));
				}
				
			}
		}

		private string GetValue(Cell cell, WorkbookPart workbookPart) {
			GetCellValue stringTypeValue = (c, w) => c.InnerText;
			GetCellValue shareStringTypeValue = (c, w) => w.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
				.ElementAt(int.Parse(c.InnerText)).InnerText;

			Dictionary<CellValues, GetCellValue> factionMap = new Dictionary<CellValues, GetCellValue>() {
				{CellValues.SharedString, shareStringTypeValue }, {CellValues.String, stringTypeValue},
				{CellValues.Number, stringTypeValue }
			};
			if(cell.DataType == null) {
				return stringTypeValue(cell, workbookPart);
			}
			return factionMap[cell.DataType.Value](cell, workbookPart);
		}


		private delegate string GetCellValue(Cell c, WorkbookPart w);

		private delegate void SetValue(WorkbookPart workbookPart);
	}

	
}
