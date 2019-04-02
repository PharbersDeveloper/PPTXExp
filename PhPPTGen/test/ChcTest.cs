using PhPPTGen.phCommand.phPpt;
using PhPPTGen.phCommand.phText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.test {
	class ChcTest {
		public void Run() {
			var input1 = new phModel.PhRequest() {
				jobid = "dcstest",
				text = new phModel.PhTextSetContent() {
					slider = 0,
					pos = new int[4] { (int)(248 / 0.000278), (int)(739 / 0.000278), (int)(2099 / 0.000278), (int)(192 / 0.000278) },
					content = "#{##[#口服降糖药市场CHC数据分析报告#C#20bule#]##P#center#}##{##[#2018Q3YTD#C#20bule#]##P#center#}#"
				}
			};
			new PhCreatePPTForXmlCommand().Exec(input1);
			new PhTextContentForXmlCommand().Exec(input1);
			var input2 = new phModel.PhRequest() {
				jobid = "dcstest",
				text = new phModel.PhTextSetContent() {
					slider = 1,
					pos = new int[4] { (int)(248 / 0.000278), (int)(739 / 0.000278), (int)(2099 / 0.000278), (int)(192 / 0.000278) },
					content = "#{##[#口服降糖药物市场规模在北京市CHC的2018Q3YTD年以47.2%的增长速度达到6.55亿人民币#C#18black#]##P#center#}#"
				}
			};
			var input3 = new phModel.PhRequest() {
				jobid = "dcstest",
				e2c = new phModel.PhExcel2Chart() {
					name = "2",
					pos = new int[4] { (int)(254 / 0.000278), (int)(432 / 0.000278), (int)(1947 / 0.000278), (int)(1284 / 0.000278) },
					chartType = "Bubble",
					slider = 3
				}
			};
			new phCommand.phChart.PhPPTImportChartCommand().Exec(input3);
		}
	}
}
