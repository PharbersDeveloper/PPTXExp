using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
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
            if (!workbookMap.ContainsKey(workbookKey)) {
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
			Console.WriteLine("Write a value to excel***********");
			Workbook workbook = new Workbook(); 
			workbookMap.TryGetValue(key, out workbook);
			Worksheet sheet = workbook.Worksheets[0];
			/**
			 * 居中需要写到css中
			 */
			//sheet.Range[p.cell].Style.VerticalAlignment = VerticalAlignType.Center;
			//sheet.Range[p.cell].Style.HorizontalAlignment = HorizontalAlignType.Center;
            foreach(string cells in p.cells) {
                string cell = new Regex("#c#[^#]+").Match(cells).Value.Replace("#c#","");
                string cate = new Regex("#t#[^#]+").Match(cells).Value.Replace("#t#", "");
                string css = new Regex("#s#[^#]+").Match(cells).Value.Replace("#s#", "");
                string value = new Regex("#v#[^#]+").Match(cells).Value.Replace("#v#", "");
                if (cell.Contains(":")) {
                    sheet.Range[cell].Merge();
                }

                if (cate == "String") {
                    sheet.Range[cell].Text = value;
                } else {
                    double tmp = 0.0;
                    if (double.TryParse(value, out tmp) && !double.IsNaN(tmp)) {
                        sheet.Range[cell].NumberFormat = "#,##0.00";
                        sheet.Range[cell].NumberValue = tmp;
                    } else {
                        sheet.Range[cell].Text = "N/A";
                    }

                }
				/**
                 * set css
                 */
				sheet.Range[cell].Style.WrapText = true;
				new PhSetXlsCssBaseCommand().Exec(cell, css, sheet);
            }
			
        }
    }
}
