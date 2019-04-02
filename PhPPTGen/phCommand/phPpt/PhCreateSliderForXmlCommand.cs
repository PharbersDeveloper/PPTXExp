using DocumentFormat.OpenXml.Packaging;
using PhPPTGen.phOpenxml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phPpt {
	class PhCreateSliderForXmlCommand : PhCommand {

		public override object Exec(params object[] parameters) {
			Console.WriteLine("PPTGen PhCommand: Insert table into pptx");
			var req = (phModel.PhRequest)parameters[0];
			if (req.slider.slider == 0) return null;
			var jobid = req.jobid;

			/**
             * 1. go to the work place
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;

			var ppt_path = workingPath + "\\" + "result.pptx";

			using (PresentationDocument pptDoc = PresentationDocument.Open(ppt_path, true)) {
				PhOpenxmlPPTHandler.GetInstance().InsertNewSlide(pptDoc, req.slider.slider, req.slider.title, req.slider.sliderType);
			}

			return null;
		}

	}
}
