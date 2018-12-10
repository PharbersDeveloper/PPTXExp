using PhPPTGen.phModel;
using Spire.Xls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phExcel {
	class PhSetXlsCssBaseCommand : PhCommand {
		private PhExcelCss css { get; set; }
		private Worksheet Sheet { get; set; }

		public override object Exec(params object[] parameters) {
			css = (PhExcelCss)parameters[0];
			Sheet = (Worksheet)parameters[1];
			
			return null;
		}

		protected void SetCss() {
			SetFontSize();
			SetFontColor();
			SetFontName();
			SetFontStyle();
			SetCellColor();
			SetCellBorders();
		}

		protected void SetFontSize() {
			Sheet.Range[css.cell].Style.Font.Size = int.Parse(css.fontSize);
		}

		protected void SetFontColor() {
			Sheet.Range[css.cell].Style.Font.Color = System.Drawing.ColorTranslator.FromHtml(css.fontColor);
		}

		protected void SetFontName() {
			Sheet.Range[css.cell].Style.Font.FontName = css.fontName;
		}

		protected void SetFontStyle() {
			Sheet.Range[css.cell].Style.Font.IsBold = ((IList)css.fontStyle).Contains("bold ");
			Sheet.Range[css.cell].Style.Font.IsItalic = ((IList)css.fontStyle).Contains("italic ");
		}

		protected void SetCellColor() {
			Sheet.Range[css.cell].Style.Color = System.Drawing.ColorTranslator.FromHtml(css.cellColor);
		}

		protected void SetCellBorders() {
			foreach(string edge in css.cellBorders) {
				switch (edge.Split('#')[0]) {
					case "top":
						Sheet.Range[css.cell].Borders[BordersLineType.EdgeTop].LineStyle = 
							(LineStyleType)Enum.Parse(typeof(LineStyleType), edge.Split('#')[1]);
						break;
					case "bottom":
						Sheet.Range[css.cell].Borders[BordersLineType.EdgeBottom].LineStyle = 
							(LineStyleType)Enum.Parse(typeof(LineStyleType), edge.Split('#')[1]);
						break;
					case "left":
						Sheet.Range[css.cell].Borders[BordersLineType.EdgeLeft].LineStyle = 
							(LineStyleType)Enum.Parse(typeof(LineStyleType), edge.Split('#')[1]);
						break;
					case "right":
						Sheet.Range[css.cell].Borders[BordersLineType.EdgeRight].LineStyle = 
							(LineStyleType)Enum.Parse(typeof(LineStyleType), edge.Split('#')[1]);
						break;
				}
			}
		}
	}
}
