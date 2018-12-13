using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using PhPPTGen.phModel;
using Spire.Xls;

namespace PhPPTGen.phCommand.phExcel {
    public class PhUpdateXlsCommand : PhCommand {
        public PhUpdateXlsCommand() {

        }

		public static Dictionary<string, Workbook> workbookMap = new Dictionary<string, Workbook>();

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
            var file_path = workingPath + "\\" + excel_name + ".xls";
			string workbookKey = jobid + excel_name;

            /**
             * 2.1 check excel is created
             *     if no create it
             */
            if (!workbookMap.ContainsKey(workbookKey) {
                CreateXlsInMap(workbookKey);
            }

            /**
             * 3. update the value in the excel
             */
            UpdateXlsInPath(workbookKey, req.push);

            return null;
        }

        private void CreateXlsInMap(string key) {
            Console.WriteLine("workbook not exist, should create one");
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
			workbookMap.Add(key, workbook);

		}

        private void UpdateXlsInPath(string key, PhExcelPush p) {
			Console.WriteLine("Write a value to excel");
			Workbook workbook = new Workbook(); 
			workbookMap.TryGetValue(key, out workbook);
			Worksheet sheet = workbook.Worksheets[0];
            if (p.cell.Contains(":")) {
                sheet.Range[p.cell].Merge();
            }

            if (p.cate == "String") {
                sheet.Range[p.cell].Text = p.value;
            } else {
                double tmp = 0.0;
                double.TryParse(p.value, out tmp);
                sheet.Range[p.cell].NumberValue = tmp;
            }
			/**
			 * set css
			 */
			phCommandFactory.PhCommandFactory fct = phCommandFactory.PhCommandFactory.GetInstance();
			fct.CreateCommandInstance(p.css.factory, p.css, sheet);
        }
    }
}
