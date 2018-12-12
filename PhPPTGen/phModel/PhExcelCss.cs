using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phModel {
	public class PhExcelCss {
		public string id { get; set; }
		public string factory { get; set; } = "PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand";
		public string cell { get; set; } = "A1";
		public string fontSize { get; set; } = "10";
		public string fontColor { get; set; } = "#000000";
		public string fontName { get; set; } = "Tahoma";
		public string[] fontStyle { get; set; } = new string[0];
		public string cellColor { get; set; } = "#FFFFFF";
		public string[] cellBorders { get; set; } = new string[0];
		public string cellBordersColor { get; set; } = "#FFFFFF";
		public string cellHeight { get; set; } = "0";
		public string cellWidth { get; set; } = "0";

	}
}
