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

		public List<List<string>> SeriesForIndex { get; } = new List<List<string>>();
		public List<List<string>> Series = new List<List<string>>();
		public List<string> CategoryLabels = new List<string>();
		public List<string> SeriesLabels = new List<string>();
		public List<List<string>> DataLabels = new List<List<string>>();
		
		public PhChartContent() { }

		public PhChartContent(List<List<string>> seriesForIndex, List<List<string>> series, List<string> categoryLabels, List<string> seriesLabels, List<List<string>> dataLabels) {
			this.SeriesForIndex = seriesForIndex;
			this.Series = series;
			this.CategoryLabels = categoryLabels;
			this.SeriesLabels = seriesLabels;
			this.DataLabels = dataLabels;
		}

		public void SetValueFromExcel(WorkbookPart workbookPart, JToken format) {
			Dictionary<string, SetValue> funcMap = new Dictionary<string, SetValue>() {
				{"row", SetValueForRowType },{"column", SetValueForColumnType},{"chcStacked", SetValueForChcStackedType}
			};

			funcMap[(string)format["contentType"]](workbookPart);
		}
		
		public string GetTitle(string titleTye) {
		Dictionary<string, GetTitleValue> titleMap = new Dictionary<string, GetTitleValue>() {
			{ "xTitle", () => SeriesLabels[0] }, { "yTitle", () => SeriesLabels[1] }
		};

			return titleMap[titleTye]();
		}

		private void SetValueForRowType(WorkbookPart workbookPart) {
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);
			SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
			var rows = sheetData.Elements<Row>().ToList();
			foreach (Cell c in rows.First()) {
				SeriesLabels.Add(GetValue(c, workbookPart));
			}

			rows.Remove(rows.First());

			foreach (Row r in rows) {
				var cells = r.Elements<Cell>().ToList();
				CategoryLabels.Add(GetValue(cells.First(), workbookPart));
				cells.Remove(cells.First());
				List<string> serise = new List<string>();
				foreach (Cell c in cells) {
					double.TryParse(GetValue(c, workbookPart), out double re);
					serise.Add(re.ToString());
				}
				Series.Add(serise);
				SeriesForIndex.Add(serise);
			}
			for(int i = 0; i < SeriesLabels.Count - Series[0].Count; i++) {
				SeriesLabels.RemoveAt(0);
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

			for (int i = 0; i < rows.First().Elements<Cell>().Count() - 1; i++) {
				Series.Add(new List<string>());
			}
			foreach (Row r in rows) {
				var cells = r.Elements<Cell>().ToList();
				SeriesLabels.Add(GetValue(cells.First(), workbookPart));
				cells.Remove(cells.First());
				List<string> serise = new List<string>();
				foreach (Cell c in cells) {
					double.TryParse(GetValue(c, workbookPart), out double re);
					Series[cells.IndexOf(c)].Add(re.ToString());
				}

			}
			foreach(List<string> s in Series) {
				SeriesForIndex.Add(s);
			}
			for (int i = 0; i < CategoryLabels.Count - Series.Count; i++) {
				CategoryLabels.RemoveAt(0);
			}
		}

		private void SetValueForChcStackedType(WorkbookPart workbookPart) {
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);
			SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
			var rows = sheetData.Elements<Row>().ToList();
			foreach (Cell c in rows.First()) {
				SeriesLabels.Add(GetValue(c, workbookPart));
			}

			rows.Remove(rows.First());

			foreach (Row r in rows) {
				var cells = r.Elements<Cell>().ToList();
				CategoryLabels.Add(GetValue(cells.First(), workbookPart));
				cells.Remove(cells.First());
				List<string> serise = new List<string>();
				foreach (Cell c in cells.Take(cells.Count / 2)) {
					double.TryParse(GetValue(c, workbookPart), out double re);
					serise.Add(re.ToString());
				}

				List<string> dataLabel = new List<string>();
				foreach (Cell c in cells.Skip(cells.Count / 2).Take(cells.Count / 2)) {
					double.TryParse(GetValue(c, workbookPart), out double re);
					dataLabel.Add(string.Format("{0:0.00%}", re));
				}
				Series.Add(serise);
				DataLabels.Add(dataLabel);
			}
			SeriesLabels.RemoveAll(x => x.Trim() == "");
			for (int i = 0; i < SeriesLabels.Count - Series[0].Count; i++) {
				SeriesLabels.RemoveAt(0);
			}
			foreach (List<string> s in Series) {
				SeriesForIndex.Add(s);
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

		private delegate string GetTitleValue();
	}

	
}
