using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;
using PhPPTGen.phModel;

namespace PhPPTGen.phOpenxml {
	class PhExcelHandler {
		private static PhExcelHandler _instance = null;

		private PhExcelHandler() { }

		public static PhExcelHandler GetInstance() {
			if (_instance == null) {
				_instance = new PhExcelHandler();
			}
			return _instance;
		}

		public void CreatExcel(string path) {
			SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
			WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
			workbookPart.Workbook = new Workbook();

			//workbookPart.AddNewPart<WorkbookStylesPart>();
			//workbookPart.WorkbookStylesPart.Stylesheet = creatStylesheet();
			//PushStyleToExcel();
			//workbookPart.WorkbookStylesPart.Stylesheet.Save();

			WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet();

			worksheetPart.Worksheet.Append(new SheetData());
			Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
			Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "pharbers" };
			sheets.Append(sheet);
			worksheetPart.Worksheet.Save();
			workbookPart.Workbook.Save();
			spreadsheetDocument.Close();
		}

		public void UpdateExcel(string path, PhExcelPush p) {
			Console.WriteLine("Write a value to excel***********");
			foreach (string cells in p.cells) {
				using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true)) {
					var iter = spreadsheetDocument.WorkbookPart.WorksheetParts.GetEnumerator();
					iter.MoveNext();
					WorksheetPart worksheetPart = iter.Current;
					//SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

					string cellReference = new Regex("#c#[^#]+").Match(cells).Value.Replace("#c#", "");
					string cate = new Regex("#t#[^#]+").Match(cells).Value.Replace("#t#", "");
					string css = new Regex("#s#[^#]+").Match(cells).Value.Replace("#s#", "");
					string value = new Regex("#v#[^#]+").Match(cells).Value.Replace("#v#", "");
					Cell cell = InsertCellIntoexcel(GetColumnName(cellReference), GetRowIndex(cellReference), worksheetPart);
					cell.CellValue = new CellValue(value);
					cell.DataType = (CellValues)Enum.Parse(typeof(CellValues), cate);

					worksheetPart.Worksheet.Save();
				}
			}
		}

		private Cell InsertCellIntoexcel(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
			
			Worksheet worksheet = worksheetPart.Worksheet;
			SheetData sheetData = worksheet.GetFirstChild<SheetData>();
			string cellReference = columnName + rowIndex;
			Row row = null;
			if (sheetData.Elements<Row>().Where(r => r.RowIndex != null && rowIndex.Equals(r.RowIndex)).Count() != 0) {
				row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
				row.CustomHeight = true;
			} else {
				row = new Row() { RowIndex = rowIndex};
				row.CustomHeight = true;
				sheetData.Append(row);
			}

			if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() != 0) {
				return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
			}

			Cell cell = null;
			foreach (Cell c in row.Elements<Cell>()) {
				if (c.CellReference.Value.Length == cellReference.Length) {
					if (String.Compare(c.CellReference.Value, cellReference, true) > 0) {
						cell = c;
						break;
					}
				}
			}

			Cell newCell = new Cell() { CellReference = cellReference };
			row.InsertBefore(newCell, cell);

			worksheet.Save();
			return newCell;
		}

		private string GetColumnName(string cellName) {
			// Create a regular expression to match the column name portion of the cell name.
			Regex regex = new Regex("[A-Za-z]+");
			Match match = regex.Match(cellName);

			return match.Value;
		}

		private uint GetRowIndex(string cellName) {
			// Create a regular expression to match the row index portion the cell name.
			Regex regex = new Regex(@"\d+");
			Match match = regex.Match(cellName);

			return uint.Parse(match.Value);
		}

		private Stylesheet CreatStylesheet() {
			Stylesheet stylesheet = new Stylesheet();
			//创建默认格式
			Fonts fonts = new Fonts();
			Font font = new Font() { FontName = new FontName { Val = "Arial" }, FontSize = new FontSize { Val = 9 } };
			fonts.Append(font);
			stylesheet.Append(fonts);
			return stylesheet;
		}

		//static void Main(string[] args) {
		//	GetInstance().CreatExcel(@"D:\alfredyang\test.xlsx");
		//	PhExcelPush p = new PhExcelPush() { cells = new string[2] { "#c#A1#t#String#v#1", "#c#B1#t#String#v#1" } };
		//	GetInstance().UpdateExcel(@"D:\alfredyang\test.xlsx", p);
		//}
	}
}
