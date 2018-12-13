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
			SetCss();
			return null;
		}

		protected void SetCss() {
			SetFontSize();
			SetFontColor();
			SetFontName();
			SetFontStyle();
			SetCellColor();
            SetCellBorders();
            SetCellBordersColor();
            SetHeight();
            SetWidth();
			SetAlignment();


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
			Sheet.Range[css.cell].Style.Font.IsBold = ((IList)css.fontStyle).Contains("bold");
			Sheet.Range[css.cell].Style.Font.IsItalic = ((IList)css.fontStyle).Contains("italic");
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

		protected void SetCellBordersColor() {
			Sheet.Range[css.cell].Borders.Color = System.Drawing.ColorTranslator.FromHtml(css.cellBordersColor);
			//Sheet.Range[css.cell].Borders.Color = System.Drawing.Color.Gainsboro;
		}

		protected void SetHeight() {
            if (!css.height.Equals("0")) Sheet.Range[css.cell].RowHeight = double.Parse(css.height);
            //if (!css.height.Equals("0")) Sheet.SetRowHeight(Sheet.Range[css.cell].RowCount, double.Parse(css.height));
        }

        protected void SetWidth() {
            if (!css.width.Equals("0")) Sheet.Range[css.cell].ColumnWidth = double.Parse(css.width);
            //if (!css.width.Equals("0")) Sheet.SetColumnWidth(Sheet.Range[css.cell].ColumnCount, double.Parse(css.width));
        }

		protected void SetAlignment() {
			
			Sheet.Range[css.cell].Style.VerticalAlignment =
				(VerticalAlignType)Enum.Parse(typeof(VerticalAlignType), css.verticalAlignType);
			Sheet.Range[css.cell].Style.HorizontalAlignment =
				(HorizontalAlignType)Enum.Parse(typeof(HorizontalAlignType), css.horizontalAlignType);
		}

	}
}
