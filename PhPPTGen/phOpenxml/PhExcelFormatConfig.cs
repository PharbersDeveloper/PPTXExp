using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace PhPPTGen.phOpenxml {
	class PhExcelFormatConfig {
		private static PhExcelFormatConfig _instance = null;
		readonly Dictionary<string, int> fontMap = new Dictionary<string, int>();
		readonly Dictionary<string, int> fillMap = new Dictionary<string, int>();
		readonly Dictionary<string, uint> numberingMap = new Dictionary<string, uint>();
		readonly Dictionary<string, int> borderMap = new Dictionary<string, int>();
		readonly Dictionary<string, int> cellformatMap = new Dictionary<string, int>();
		readonly Dictionary<Boolean, IPhExcelFormatHandler> FormatHandlerMap = new Dictionary<Boolean, IPhExcelFormatHandler>(); 
		readonly XmlDocument _doc = null;

		public static PhExcelFormatConfig GetInstans() {
			if (_instance == null) {
				_instance = new PhExcelFormatConfig();
			}

			return _instance;
		}

		public void OneExcelOver() {
			cellformatMap.Clear();
		}

		public int GetCellFormatIndexByName(Stylesheet ss, string name) {
			return FormatHandlerMap[cellformatMap.TryGetValue(name, out int index)].GetCellFormatId(ss, name, index);
		}

		protected PhExcelFormatConfig() {
			_doc = new XmlDocument();
			_doc.Load(PhConfigHandler.GetInstance().path + @"\PhFormatConfig.xml");
			IPhExcelFormatHandler getFormatHandler = new GetFormat();
			IPhExcelFormatHandler addFormatHandler = new AddFormat();
			FormatHandlerMap.Add(true, getFormatHandler);
			FormatHandlerMap.Add(false, addFormatHandler);
		}

		public void PushCellFormatsToStylesheet(Stylesheet ss) {
			PushFont(ss);
			PushtNumbering(ss);
			PushBorder(ss);
			PushFill(ss);
			PushCellFormat(ss);
		}

		public void AddCellFormat(Stylesheet ss, string name) {
			CellFormats cfs = ss.CellFormats;
			var id = name;

			var fontId = new Regex(@"\*font\*[^\*]*").Match(name).Value.Replace("*font*", "");
			//int fontIdx = fontMap[""];
			fontMap.TryGetValue(fontId, out int fontIdx);

			var fillId = new Regex(@"\*fill\*[^\*]*").Match(name).Value.Replace("*fill*", "");
			//var fillIdx = fillMap[""];
			fillMap.TryGetValue(fillId, out int fillIdx);

			var numberingId = new Regex(@"\*num\*[^\*]*").Match(name).Value.Replace("*num*", "");
			//var numberingIdx = numberingMap[""];
			numberingMap.TryGetValue(numberingId, out uint numberingIdx);

			var borderId = new Regex(@"\*border\*[^\*]*").Match(name).Value.Replace("*border*", "");
			//var borderIdx = borderMap[""];
			borderMap.TryGetValue(borderId, out int borderIdx);

			var hv = (HorizontalAlignmentValues)Enum.Parse(typeof(HorizontalAlignmentValues),
				new Regex(@"\*h\*[^\*]*").Match(name).Value.Replace("*h*", ""));
			var vv = (VerticalAlignmentValues)Enum.Parse(typeof(VerticalAlignmentValues),
				new Regex(@"\*v\*[^\*]*").Match(name).Value.Replace("*v*", ""));
			var cf = new CellFormat() {
				NumberFormatId = numberingIdx,
				FontId = (uint)fontIdx,
				FillId = (uint)fillIdx,
				BorderId = (uint)borderIdx,
				Alignment = new Alignment() { Horizontal = hv, Vertical = vv },
				ApplyNumberFormat = true,
			};

			cfs.Append(cf);

			var idx = cfs.Elements<CellFormat>().Count() - 1;
			cellformatMap[id] = idx;
		}

		protected void PushFont(Stylesheet ss) {
			Fonts fonts = ss.Fonts;
			XmlNode fontsNode = _doc.SelectSingleNode("stylesheet/fonts");
			foreach (XmlNode fontNode in fontsNode.SelectNodes("font")) {
				string id = fontNode.Attributes.GetNamedItem("id").Value;
				Console.WriteLine("push font config" + id);
				string name = fontNode.Attributes.GetNamedItem("name").Value;
				string size = fontNode.Attributes.GetNamedItem("size").Value;
				string bold = fontNode.Attributes.GetNamedItem("bold").Value;
				string color = fontNode.Attributes.GetNamedItem("color").Value;
				Font font = new Font();
				font.Append(new FontName { Val = name }, new FontSize { Val = Double.Parse(size) },
					new Color { Rgb = new HexBinaryValue(color) });
				if (Boolean.Parse(bold)) {
					font.Append(new Bold());
				}

				fonts.Append(font);

				var idx = fonts.Elements<Font>().Count() - 1;
				fontMap[id] = idx;
			}
		}

		protected void PushtNumbering(Stylesheet ss) {
			var numberings = ss.NumberingFormats;

			var xn = _doc.SelectSingleNode("stylesheet/numberings");
			var nlst = xn.SelectNodes("numbering");
			Console.WriteLine(nlst.Count);

			foreach (XmlNode f in nlst) {
				string id = f.Attributes.GetNamedItem("id").Value;
				uint idx = uint.Parse(f.Attributes.GetNamedItem("idx").Value);
				string code = f.Attributes.GetNamedItem("code").Value;
				Console.WriteLine("push numbering" + id + " " + code);

				NumberingFormat nf = new NumberingFormat { NumberFormatId = idx, FormatCode = code };
				numberings.Append(nf);
				numberingMap[id] = idx;
			}
		}

		protected void PushBorder(Stylesheet ss) {
			var borders = ss.Borders;

			XmlNode borderConfig = _doc.SelectSingleNode("stylesheet/borders");
			foreach (XmlNode borderNode in borderConfig.SelectNodes("border")) {
				string id = borderNode.Attributes.GetNamedItem("id").Value;
				Border border = new Border();
				var left = borderNode.SelectSingleNode("left");
				var leftStyle = left.Attributes.GetNamedItem("style").Value;
				var leftColor = left.Attributes.GetNamedItem("color").Value;
				var lb = new LeftBorder() { Style = (BorderStyleValues)Enum.Parse(typeof(BorderStyleValues), leftStyle) };
				Color lc = new Color() { Rgb = leftColor };
				lb.Append(lc);
				border.Append(lb);

				var right = borderNode.SelectSingleNode("right");
				var rightStyle = right.Attributes.GetNamedItem("style").Value;
				var rightColor = right.Attributes.GetNamedItem("color").Value;
				var rb = new RightBorder() { Style = (BorderStyleValues)Enum.Parse(typeof(BorderStyleValues), rightStyle) };
				Color rc = new Color() { Rgb = rightColor };
				rb.Append(rc);
				border.Append(rb);

				var top = borderNode.SelectSingleNode("top");
				var topStyle = top.Attributes.GetNamedItem("style").Value;
				var topColor = top.Attributes.GetNamedItem("color").Value;
				var tb = new TopBorder() { Style = (BorderStyleValues)Enum.Parse(typeof(BorderStyleValues), topStyle) };
				Color tc = new Color() { Rgb = topColor };
				tb.Append(tc);
				border.Append(tb);

				var bottom = borderNode.SelectSingleNode("bottom");
				var bottomStyle = bottom.Attributes.GetNamedItem("style").Value;
				var bottomColor = bottom.Attributes.GetNamedItem("color").Value;
				var bb = new BottomBorder() { Style = (BorderStyleValues)Enum.Parse(typeof(BorderStyleValues), bottomStyle) };
				Color bc = new Color() { Rgb = bottomColor };
				bb.Append(bc);
				border.Append(bb);

				borders.Append(border);

				var idx = borders.Elements<Border>().Count() - 1;
				borderMap[id] = idx;

			}
		}

		protected void PushFill(Stylesheet ss) {
			Fills fills = ss.Fills;

			XmlNode xn = _doc.SelectSingleNode("stylesheet/fills");
			XmlNodeList nlst = xn.SelectNodes("fill");
			Console.WriteLine(nlst.Count);

			foreach (XmlNode f in nlst) {
				string id = f.Attributes.GetNamedItem("id").Value;
				Console.WriteLine("push fill" + id);

				string fill_type = f.Attributes.GetNamedItem("type").Value;
				string fill_color = f.Attributes.GetNamedItem("color").Value;
				Fill fill = new Fill();

				PatternFill pf = new PatternFill() { PatternType = PatternValues.Solid };
				ForegroundColor fc = new ForegroundColor() { Rgb = fill_color };
				BackgroundColor bc = new BackgroundColor() { Indexed = (UInt32Value)64U };

				pf.Append(fc);
				pf.Append(bc);
				fill.Append(pf);

				fills.Append(fill);
				int idx = fills.Elements<Fill>().Count() - 1;
				fillMap[id] = idx;
			}
		}

		private void PushCellFormat(Stylesheet ss) {
			CellFormats cfs = ss.CellFormats;

			XmlNode xn = _doc.SelectSingleNode("stylesheet/cellformats");
			XmlNodeList nlst = xn.SelectNodes("cellformat");
			Console.WriteLine(nlst.Count);

			foreach (XmlNode f in nlst) {
				var id = f.Attributes.GetNamedItem("id").Value;

				var fontId = f.Attributes.GetNamedItem("font").Value;
				var fontIdx = fontMap[fontId];

				var fillId = f.Attributes.GetNamedItem("fill").Value;
				var fillIdx = fillMap[fillId];

				var numberingId = f.Attributes.GetNamedItem("numbering").Value;
				var numberingIdx = numberingMap[numberingId];

				var borderId = f.Attributes.GetNamedItem("border").Value;
				var borderIdx = borderMap[borderId];

				var hv = (HorizontalAlignmentValues)Enum.Parse(typeof(HorizontalAlignmentValues),
					f.Attributes.GetNamedItem("horizontal").Value);
				var vv = (VerticalAlignmentValues)Enum.Parse(typeof(VerticalAlignmentValues),
					f.Attributes.GetNamedItem("vertical").Value);
				var cf = new CellFormat() {
					NumberFormatId = numberingIdx,
					FontId = (uint)fontIdx,
					FillId = (uint)fillIdx,
					BorderId = (uint)borderIdx,
					Alignment = new Alignment() { Horizontal = hv, Vertical = vv },
					ApplyNumberFormat = true,
				};

				cfs.Append(cf);

				var idx = cfs.Elements<CellFormat>().Count() - 1;
				cellformatMap[id] = idx;
			}
		}

		class GetFormat : IPhExcelFormatHandler {
			public int GetCellFormatId(Stylesheet ss, string name, int index) {
				return index;
			}
		}

		class AddFormat : IPhExcelFormatHandler {
			public int GetCellFormatId(Stylesheet ss, string name, int index) {
				PhExcelFormatConfig.GetInstans().AddCellFormat(ss, name);
				return PhExcelFormatConfig.GetInstans().GetCellFormatIndexByName(ss, name);
			}
		}

		//protected string GetAttributeValue(XmlNode node, string name) {
		//	return node.Attributes.GetNamedItem(name).Value;
		//}
	}
}
