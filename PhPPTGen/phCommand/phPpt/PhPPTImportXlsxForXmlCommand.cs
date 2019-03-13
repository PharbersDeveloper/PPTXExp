using DocumentFormat.OpenXml.Packaging;
using PhPPTGen.phOpenxml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phPpt {
	class PhPPTImportXlsxForXmlCommand: PhCommand {
		public override object Exec(params object[] parameters) {
			Console.WriteLine("PPTGen PhCommand: Insert table into pptx");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var e2p = req.e2p;

			//清除格式信息
			PhExcelFormatConfig.GetInstans().OneExcelOver();

			/**
             * 1. go to the work place
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;
			string workbookKey = jobid + e2p.name;

			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2p.name + ".xlsx";
			var emfPath = workingPath + "\\" + e2p.name + ".emf";
			if (!File.Exists(ePath)) {
				throw new Exception("Excel name is not exists");
			}

			/**
			 * 3.获得展示excle的emf
			 */
			Workbook workbook = new Workbook();
			workbook.LoadFromFile(ePath);
			Worksheet sheet = workbook.Worksheets[0];
			Stream emfstram = new FileStream(emfPath, FileMode.CreateNew);
			sheet.SaveToImage(emfstram, 1, 1, sheet.Rows.Length, sheet.Columns.Length, System.Drawing.Imaging.EmfType.EmfPlusDual);
			emfstram.Close();

			/**
			 * 4.将excel以及emf插入ppt最后一页
			 */

			var ppt_path = workingPath + "\\" + "result.pptx";

			using (PresentationDocument pptDoc = PresentationDocument.Open(ppt_path, true)) {
				//insert new ppt
				if (pptDoc.PresentationPart.Presentation.SlideIdList.Count() - 1 < e2p.slider) {
					PhOpenxmlPPTHandler.GetInstance().InsertNewSlide(pptDoc, e2p.slider, "");
				}

				PhOpenxmlPPTHandler.GetInstance().InsertExcel(pptDoc, ePath, emfPath, e2p.pos);
			}

			return null;
		}
	}
}
