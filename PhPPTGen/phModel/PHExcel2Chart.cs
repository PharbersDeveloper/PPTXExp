using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phModel {
	public class PhExcel2Chart {
		public string id { get; set; }
		public string name { get; set; }
		public string chartType { get; set; }
		public string css { get; set; }
		public int[] pos { get; set; }
		public int slider { get; set; }
	}
}
