using System;
using System.IO;
using PhPPTGen.phModel;
using Spire.Xls;

namespace PhPPTGen.phCommand.phExcel {
    public class PhUpdateXlsCommand : PhCommand {
        public PhUpdateXlsCommand() {

        }

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

            /**
             * 2.1 check excel is created
             *     if no create it
             */
            if (!File.Exists(file_path)) {
                CreateXlsInPath(file_path);
            }

            /**
             * 3. update the value in the excel
             */
            UpdateXlsInPath(file_path, req.push);

            return null;
        }

        private void CreateXlsInPath(string path) {
            Console.WriteLine("File not exist, should create one");
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A1"].Text = "Hello,World!";
            workbook.SaveToFile(path);
        }

        private void UpdateXlsInPath(string path, PhExcelPush p) {
            Console.WriteLine("Write a value to excel");
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);
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

            workbook.SaveToFile(path);
        }
    }
}
