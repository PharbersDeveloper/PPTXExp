using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phModel {
	public class PhExcelCss {
		public string id { get; set; }
		public string factory { get; set; } = "PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand";
		public string cell { get; set; }
		public string fontSize { get; set; } = "10";
		public string fontColor { get; set; } = "#000000";
		public string fontName { get; set; }
		public string[] fontStyle { get; set; }
		public string cellColor { get; set; } = "#FFFFFF";
		public string[] cellBorders { get; set; }
		public string cellBordersColor { get; set; } = "#FFFFFF";
	}
}
