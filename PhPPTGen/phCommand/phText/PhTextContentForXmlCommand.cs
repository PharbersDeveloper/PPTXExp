using DocumentFormat.OpenXml.Packaging;
using PhPPTGen.phOpenxml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phText {
	class PhTextContentForXmlCommand: PhCommand {
		public override object Exec(params object[] parameters) {
			Console.WriteLine("PPTGen phCommand: generate text shape for ppt");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var text = req.text;

			/**
             * 1. go to the tmp dir
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;
			var file_path = workingPath + "\\result.pptx";

			using (PresentationDocument pptDoc = PresentationDocument.Open(file_path, true)) {

				if (pptDoc.PresentationPart.Presentation.SlideIdList.Count() - 1 < text.slider) {
					PhOpenxmlPPTHandler.GetInstance().InsertNewSlide(pptDoc, text.slider, "");
				}
				PhOpenxmlPPTHandler.GetInstance().InsertText(pptDoc, text);
			}			
			return null;
		}
	}
}
