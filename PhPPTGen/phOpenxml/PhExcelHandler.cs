using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;
using PhPPTGen.phModel;
using PhPPTGen.phOpenxml.phFormatHandler;
using System.Collections.Generic;
using PhPPTGen.phOpenxml.phExcelChart.PhChartElement;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phOpenxml {
	class PhExcelHandler {
		private static PhExcelHandler _instance = null;
		private readonly Dictionary<String, CheckValue> valueCheckMap = new Dictionary<string, CheckValue>();
		private readonly Dictionary<string, JToken> ChartTypeMap;

		private PhExcelHandler() {
			string checkString(string v, out string type) {
				type = "String";
				return v;
			}
			string checkNumber(string v, out string type) {
				type = new Dictionary<Boolean, String> { { true, "Number" }, { false, "String" } }[Double.TryParse(v, out double re) && !Double.IsNaN(re)];
				return new Dictionary<String, String> { { "Number", v }, { "String", "N/A" } }[type];
			}

			valueCheckMap["String"] = checkString;
			valueCheckMap["Number"] = checkNumber;

			ChartTypeMap = new Dictionary<string, JToken>();
			foreach (JToken jToken in PhConfigHandler.GetInstance().configMap["chartType"]) {
				using (StreamReader reader = File.OpenText(PhConfigHandler.GetInstance().path + jToken.First().Value<string>())) {
					ChartTypeMap.Add(((JProperty)jToken).Name, JToken.ReadFrom(new JsonTextReader(reader)));
				}
			}
		}

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

			workbookPart.AddNewPart<WorkbookStylesPart>();
			workbookPart.WorkbookStylesPart.Stylesheet = CreateStylesheet();
			PhExcelFormatConfig.GetInstans().PushCellFormatsToStylesheet(workbookPart.WorkbookStylesPart.Stylesheet);
			workbookPart.WorkbookStylesPart.Stylesheet.Save();

			WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet();

			Columns columns = new Columns();
			columns.Append(new Column { Min = 1, Max = 1, Width = 10, CustomWidth = true });
			worksheetPart.Worksheet.Append(columns);

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
					value = valueCheckMap[cate](value, out cate);
					Cell cell = InsertCellIntoexcel(GetColumnName(cellReference), GetRowIndex(cellReference), worksheetPart);
					cell.CellValue = new CellValue(value);
					cell.DataType = (CellValues)Enum.Parse(typeof(CellValues), cate);
					var excelCss = PhExcelCssHandler.GetInstance().Css2CellFormatName(css);
					cell.StyleIndex = (uint)PhExcelFormatConfig.GetInstans().GetCellFormatIndexByName
						(spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet, GetCellFormatName(excelCss));
					MergeCell(worksheetPart, cellReference.Split(':').First(), cellReference.Split(':').Last());
					SetRowHeight(Double.Parse(excelCss.height), GetRowIndex(cellReference), worksheetPart);
					SetColWidth(Double.Parse(excelCss.width), GetColumnName(cellReference), worksheetPart);
					worksheetPart.Worksheet.Save();
				}
			}
		}

		public void InsertChartIntoExcel(WorkbookPart workbookPart, string type) {
			PhChartPartsHandler handler = new PhChartPartsHandler {
				Format = ChartTypeMap[type]["chart"]
			};
			handler.Content.SetValueFromExcel(workbookPart, handler.Format);
			WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(0);
			DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
			worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
			worksheetPart.Worksheet.Save();
			ChartPart chartPart1 = drawingsPart.AddNewPart<ChartPart>("ch1");
			handler.CreateChartPart(chartPart1);
			CreateDrawingPart(drawingsPart);
			chartPart1.ChartSpace.Save();
			drawingsPart.WorksheetDrawing.Save();

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
				row = new Row() { RowIndex = rowIndex };
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

		private Stylesheet CreateStylesheet() {
			var ss = new Stylesheet();

			var fts = new Fonts();
			var ftn = new FontName { Val = "Arial" };
			var ftsz = new FontSize { Val = 11 };
			var ft = new DocumentFormat.OpenXml.Spreadsheet.Font { FontName = ftn, FontSize = ftsz };
			fts.Append(ft);
			fts.Count = (uint)fts.ChildElements.Count;

			var fills = new Fills();
			var fill = new Fill();
			var patternFill = new PatternFill { PatternType = PatternValues.None };
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fill = new Fill();
			patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fills.Count = (uint)fills.ChildElements.Count;

			var borders = new Borders();
			var border = new Border {
				LeftBorder = new LeftBorder(),
				RightBorder = new RightBorder(),
				TopBorder = new TopBorder(),
				BottomBorder = new BottomBorder(),
				DiagonalBorder = new DiagonalBorder()
			};
			borders.Append(border);
			borders.Count = (uint)borders.ChildElements.Count;

			var csfs = new CellStyleFormats();
			var cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 };
			csfs.Append(cf);
			csfs.Count = (uint)csfs.ChildElements.Count;

			// dd/mm/yyyy is also Excel style index 14

			uint iExcelIndex = 164;
			var nfs = new NumberingFormats();
			var cfs = new CellFormats();

			cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
			cfs.Append(cf);

			var nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "dd/mm/yyyy hh:mm:ss" };
			nfs.Append(nf);

			cf = new CellFormat {
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			iExcelIndex = 165;
			nfs = new NumberingFormats();
			cfs = new CellFormats();

			cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
			cfs.Append(cf);

			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "MMM yyyy" };
			nfs.Append(nf);

			cf = new CellFormat {
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			iExcelIndex = 170;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.0000" };
			nfs.Append(nf);
			cf = new CellFormat {
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			// #,##0.00 is also Excel style index 4
			iExcelIndex = 171;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.00" };
			nfs.Append(nf);
			cf = new CellFormat {
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			// @ is also Excel style index 49
			iExcelIndex = 172;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "@" };
			nfs.Append(nf);
			cf = new CellFormat {
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			nfs.Count = (uint)nfs.ChildElements.Count;
			cfs.Count = (uint)cfs.ChildElements.Count;

			ss.Append(nfs);
			ss.Append(fts);
			ss.Append(fills);
			ss.Append(borders);
			ss.Append(csfs);
			ss.Append(cfs);

			var css = new CellStyles();
			var cs = new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 };
			css.Append(cs);
			css.Count = (uint)css.ChildElements.Count;
			ss.Append(css);

			var dfs = new DifferentialFormats { Count = 0 };
			ss.Append(dfs);

			var tss = new TableStyles {
				Count = 0,
				DefaultTableStyle = "TableStyleMedium9",
				DefaultPivotStyle = "PivotStyleLight16"
			};
			ss.Append(tss);
			return ss;
		}

		private string GetCellFormatName(PhExcelCssForOpenxml css) {
			return "*font*" + css.fontName + css.fontSize + css.bold + css.fontColor + "*fill*" + css.cellColor + "*num*" + css.numbering
				+ "*border*" + css.topBorder + css.bottomBorder + css.leftBorder + css.rightBorder
				+ "*h*" + css.horizontalAlignType + "*v*" + css.verticalAlignType;
		}

		private void MergeCell(WorksheetPart workSheetPart, string c1, string c2) {
			var worksheet = workSheetPart.Worksheet;
			SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();
			if (c1 == c2) {
				return;
			}
			//InsertCellIntoexcel(GetColumnName(c1), GetRowIndex(c2), workSheetPart);
			//InsertCellIntoexcel(GetColumnName(c1), GetRowIndex(c2), workSheetPart);

			MergeCells mergeCells;
			if (worksheet.Elements<MergeCells>().Count() > 0) {
				mergeCells = worksheet.Elements<MergeCells>().First();
			} else {
				mergeCells = new MergeCells();

				// Insert a MergeCells object into the specified position.
				if (worksheet.Elements<CustomSheetView>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
				} else if (worksheet.Elements<DataConsolidate>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
				} else if (worksheet.Elements<SortState>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
				} else if (worksheet.Elements<AutoFilter>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
				} else if (worksheet.Elements<Scenarios>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
				} else if (worksheet.Elements<ProtectedRanges>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
				} else if (worksheet.Elements<SheetProtection>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
				} else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0) {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
				} else {
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
				}
			}

			// Create the merged cell and append it to the MergeCells collection.
			MergeCell mergeCell = new MergeCell() { Reference = new StringValue(c1 + ":" + c2) };
			mergeCells.Append(mergeCell);

			worksheet.Save();

		}

		private void SetRowHeight(double height, uint rowIndex, WorksheetPart worksheetPart) {
			Worksheet worksheet = worksheetPart.Worksheet;
			SheetData sheetData = worksheet.GetFirstChild<SheetData>();
			sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First().Height = new DoubleValue(height); 
		}

		private void SetColWidth(double width, string columnName, WorksheetPart worksheetPart) {
			Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
			//Columns columns = new Columns();
			uint columnIndex = (uint)(columnName.First() - 'A' + 1);
			columns.Append(new Column {
				Min = new UInt32Value(columnIndex),
				Max = new UInt32Value(columnIndex),
				Width = new DoubleValue(width),
				CustomWidth = true
			});
			//worksheetPart.Worksheet.Append(columns);
		}

		private delegate string CheckValue(string value, out string type);

		private static void CreateDrawingPart(DrawingsPart drawingsPart) {
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

		//static void Main(string[] args) {
		//	GetInstance().CreatExcel(@"D:\alfredyang\test.xlsx");
		//	GetInstance().CreatExcel(@"D:\alfredyang\test2.xlsx");
		//	PhExcelPush p = new PhExcelPush() {
		//		cells = new string[3] { "#c#A1#t#Number#v#big#s#row_title_common*row_7", "#c#B1#t#Number#v#1.3#s#col_common3*col_title_common",
		//		"#c#A2#t#Number#v#1.123#s#row_title_common*row_7"}
		//	};
		//	GetInstance().UpdateExcel(@"D:\alfredyang\test.xlsx", p);
		//	PhExcelFormatConfig.GetInstans().OneExcelOver();
		//	GetInstance().UpdateExcel(@"D:\alfredyang\test2.xlsx", p);
		//}
	}
}
