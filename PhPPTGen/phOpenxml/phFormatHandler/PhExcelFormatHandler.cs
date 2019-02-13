using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phOpenxml {
	interface IPhExcelFormatHandler {
		int GetCellFormatId(Stylesheet ss, string name, int index);

	}
}
