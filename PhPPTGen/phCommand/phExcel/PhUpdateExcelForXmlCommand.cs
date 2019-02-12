using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PhPPTGen.phOpenxml;

namespace PhPPTGen.phCommand.phExcel {
	class PhUpdateExcelForXmlCommand: PhCommand {
		public override object Exec(params object[] parameters) {
			Console.WriteLine("Execute Commmand: PhUpdateXls update value command");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;

			/**
             * 1. query temp dir
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;

			/**
             * 2. query excel xls file in the working dir
             */
			var excel_name = req.push.name;
			Console.WriteLine("push Value to Excel");
			Console.WriteLine(excel_name);
			var file_path = workingPath + "\\" + excel_name + ".xlsx";
			string workbookKey = jobid + excel_name;

			/**
             * 2.1 check excel is created
             *     if no create it
             */
			if (!File.Exists(file_path)) {
				PhExcelHandler.GetInstance().CreatExcel(file_path);
			}

			/**
             * 3. update the value in the excel
             */
			PhExcelHandler.GetInstance().UpdateExcel(file_path, req.push);

			return null;
		}
	}
}
