using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phModel {
	public class PhExcelCss {
		public string factory { get; set; }
		public string cell { get; set; } 
		public string fontSize { get; set; } 
		public string fontColor { get; set; } 
		public string fontName { get; set; }
        public string[] fontStyle { get; set; } = new string[0];
		public string cellColor { get; set; }
        public string[] cellBorders { get; set; } = new string[0];
		public string cellBordersColor { get; set; } 
		public string height { get; set; } 
		public string width { get; set; } 
		public string verticalAlignType { get; set; } 
		public string horizontalAlignType { get; set; } 

	}
}
